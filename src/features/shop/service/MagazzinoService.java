package features.shop.service;

import core.db.DbManager;
import features.shop.model.Articolo;
import features.shop.model.Riga;
import features.shop.ui.ShopAppFrame;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.*;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

public class MagazzinoService {
    private final DbManager dbManager;
    private final Properties props;
    private final java.io.File configFile; // non usata qui, ma utile se in futuro vuoi salvare

    public MagazzinoService(DbManager dbManager, Properties props, java.io.File configFile) {
        this.dbManager = dbManager;
        this.props = props;
        this.configFile = configFile;
    }

    public void exportMagazzino(JFrame parent, String outPath, String anagPath, ShopAppFrame gui) {
        if (dbManager == null || outPath == null || outPath.isEmpty() || anagPath == null || anagPath.isEmpty()) {
            JOptionPane.showMessageDialog(parent,
                    "Controlla le impostazioni di connessione Firebird, il file di output e il file di raffronto!",
                    "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        gui.setStatButtonEnabled(false);
        gui.setProgressIndeterminate(true);
        gui.setProgressBarVisible(true);
        gui.setStatus("Elaborazione magazzino in corso...");

        new SwingWorker<Void, Integer>() {
            @Override
            protected Void doInBackground() {
                try {
                    runElaborazione(outPath, anagPath, this::publish, gui);
                } catch (Exception ex) {
                    gui.appendLog("[ERRORE] (Magazzino) " + ex);
                    SwingUtilities.invokeLater(() ->
                            JOptionPane.showMessageDialog(parent, "Errore: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE));
                }
                return null;
            }

            @Override
            protected void process(java.util.List<Integer> chunks) {
                int last = chunks.get(chunks.size() - 1);
                gui.setProgressIndeterminate(false);
                gui.setProgress(last);
                gui.setStatus("Avanzamento esportazione magazzino: " + last + "%");
            }

            @Override
            protected void done() {
                gui.setProgressBarVisible(false);
                gui.setStatButtonEnabled(true);
                gui.setStatus("Esportazione magazzino completata!");
                gui.appendLog("[LOG] Esportazione magazzino completata.");
                JOptionPane.showMessageDialog(parent, "Esportazione magazzino completata!", "Esito", JOptionPane.INFORMATION_MESSAGE);
            }
        }.execute();
    }

    public void runElaborazione(String outPath, String anagPath, Consumer<Integer> progressCallback, ShopAppFrame gui) throws Exception {
        IOUtils.setByteArrayMaxOverride(200_000_000);

        gui.appendLog("[LOG] Inizio elaborazione");
        progressCallback.accept(2);

        Map<String, Articolo> byBARCODE = new HashMap<>();
        gui.appendLog("[LOG] Apro file anagrafica: " + anagPath);

        try (FileInputStream fis = new FileInputStream(anagPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            Map<String, Integer> colIndex = new HashMap<>();
            Row header = sheet.getRow(0);
            for (Cell cell : header) {
                String v = cell.getStringCellValue().trim().toUpperCase();
                colIndex.put(v, cell.getColumnIndex());
            }

            int anagCount = 0;
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String barcode = getCellString(row, colIndex.getOrDefault("BARCODE", -1));
                if (barcode.isEmpty() || !(barcode.startsWith("0") || barcode.startsWith("8"))) continue;

                String modello = getCellString(row, colIndex.getOrDefault("MODELLO", -1));
                String tessuto = getCellString(row, colIndex.getOrDefault("TESSUTO", -1));
                String trattamento = getCellString(row, colIndex.getOrDefault("TRATTAMENTO", -1));
                String colore = getCellString(row, colIndex.getOrDefault("COLORE", -1));
                String taglia = getCellString(row, colIndex.getOrDefault("TAGLIA", -1));
                String genere = getCellString(row, colIndex.getOrDefault("GENERE", -1));
                String tipologia = getCellString(row, colIndex.getOrDefault("TIPOLOGIA", -1));

                Articolo art = new Articolo(modello, tessuto, trattamento, colore, taglia, genere, tipologia);
                byBARCODE.put(barcode, art);
                anagCount++;
            }

            gui.appendLog("[LOG] Righe anagrafica caricate: " + anagCount);
            gui.appendLog("[LOG] Chiavi byBARCODE: " + byBARCODE.size());
        }

        progressCallback.accept(10);

        List<Riga> righe = new ArrayList<>();
        int dbCount = 0, matchCount = 0;

        gui.appendLog("[LOG] Connetto a DB");
        try (Connection conn = dbManager.getConnection()) {
            String sql = """
                SELECT "EAN",
                       CAST(COALESCE(SUM("Caricato"), 0) - COALESCE(SUM("Scaricato"), 0) AS INTEGER) AS "Giacenza"
                FROM (
                    SELECT "TArticoliTaglieColori"."CodBarre" AS "EAN",
                           "TMovMagazz"."QtaCaricata" AS "Caricato",
                           "TMovMagazz"."QtaScaricata" AS "Scaricato"
                    FROM "TMovMagazz"
                    LEFT OUTER JOIN "TArticoli"
                        ON ("TArticoli"."IDArticolo" = "TMovMagazz"."IDArticolo")
                    LEFT OUTER JOIN "TArticoliTaglieColori"
                        ON (
                            "TArticoliTaglieColori"."IDArticolo" = "TMovMagazz"."IDArticolo"
                            AND "TArticoliTaglieColori"."Taglia" = SUBSTRING("TMovMagazz"."Lotto" FROM 1 FOR POSITION('/', "TMovMagazz"."Lotto") - 1)
                            AND "TArticoliTaglieColori"."Colore" = SUBSTRING("TMovMagazz"."Lotto" FROM POSITION('/', "TMovMagazz"."Lotto") + 1)
                        )
                )
                GROUP BY "EAN"
            """;

            try (Statement stmt = conn.createStatement();
                 ResultSet rs = stmt.executeQuery(sql)) {
                while (rs.next()) {
                    String barcode = rs.getString("EAN");
                    int giacenza = rs.getInt("Giacenza");
                    dbCount++;

                    Articolo art = byBARCODE.get(barcode);
                    if (art != null) {
                        righe.add(new Riga(
                                art.modello, art.tessuto, art.trattamento, art.colore, art.taglia,
                                art.genere, art.tipologia, giacenza
                        ));
                        matchCount++;
                    }
                }
            }
        }

        gui.appendLog("[LOG] Righe trovate nel DB: " + dbCount);
        gui.appendLog("[LOG] Righe abbinate all'anagrafica: " + matchCount);

        List<Riga> uomoPantalone = new ArrayList<>();
        List<Riga> donnaPantalone = new ArrayList<>();
        List<Riga> altro = new ArrayList<>();

        for (Riga r : righe) {
            String gNorm = (r.genere == null ? "" : r.genere.trim().toUpperCase());
            String tNorm = (r.tipologia == null ? "" : r.tipologia.trim().toUpperCase());
            if (gNorm.startsWith("UOMO") && tNorm.startsWith("PANT")) uomoPantalone.add(r);
            else if (gNorm.startsWith("DONNA") && tNorm.startsWith("PANT")) donnaPantalone.add(r);
            else altro.add(r);
        }

        List<String> taglieUomo = collectTaglie(uomoPantalone);
        List<String> taglieDonna = collectTaglie(donnaPantalone);
        List<String> taglieAltro = collectTaglie(altro);

        Comparator<Riga> comparator = Comparator
                .comparing((Riga r) -> r.modello == null ? "" : r.modello)
                .thenComparing(r -> r.tessuto == null ? "" : r.tessuto)
                .thenComparing(r -> r.trattamento == null ? "" : r.trattamento)
                .thenComparing(r -> r.colore == null ? "" : r.colore);

        uomoPantalone.sort(comparator);
        donnaPantalone.sort(comparator);
        altro.sort(comparator);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            scriviFoglioPivotConStile(wb, "UOMO_PANTALONE", uomoPantalone, taglieUomo);
            scriviFoglioPivotConStile(wb, "DONNA_PANTALONE", donnaPantalone, taglieDonna);
            scriviFoglioPivotConStile(wb, "ALTRO", altro, taglieAltro);

            try (FileOutputStream out = new FileOutputStream(outPath)) {
                wb.write(out);
            }
        }

        progressCallback.accept(100);
        gui.appendLog("[LOG] Esportazione magazzino completata!");
    }

    private static List<String> collectTaglie(List<Riga> righe) {
        Set<String> taglieSet = new TreeSet<>((a, b) -> {
            try { return Integer.compare(Integer.parseInt(a.trim()), Integer.parseInt(b.trim())); }
            catch (Exception e) { return a.trim().compareTo(b.trim()); }
        });

        for (Riga r : righe) {
            String taglia = r.taglia;
            if (taglia != null && !taglia.trim().isEmpty()) taglieSet.add(taglia.trim());
        }
        return new ArrayList<>(taglieSet);
    }

    private static void scriviFoglioPivotConStile(XSSFWorkbook wb, String nomeFoglio, List<Riga> righe, List<String> taglie) {
        if (righe.isEmpty()) return;

        String safeName = nomeFoglio.replaceAll("[^A-Za-z0-9_]", "");
        if (safeName.length() > 31) safeName = safeName.substring(0, 31);
        XSSFSheet sheet = wb.createSheet(safeName);

        XSSFCellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont headerFont = wb.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        headerStyle.setBorderTop(BorderStyle.THICK);
        headerStyle.setBorderBottom(BorderStyle.THICK);
        headerStyle.setBorderLeft(BorderStyle.THICK);
        headerStyle.setBorderRight(BorderStyle.THICK);
        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        XSSFCellStyle dataStyle = wb.createCellStyle();
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        int colTotMagazzino = 4 + taglie.size() + 2;
        Row magazzinoRow = sheet.createRow(0);
        Cell cellLabel = magazzinoRow.createCell(colTotMagazzino);
        cellLabel.setCellValue("TOTALE MAGAZZINO");
        cellLabel.setCellStyle(headerStyle);

        Row header = sheet.createRow(1);
        String[] intest = {"MODELLO", "TESSUTO", "TRATTAMENTO", "COLORE"};
        for (int i = 0; i < intest.length; i++) {
            Cell c = header.createCell(i);
            c.setCellValue(intest[i]);
            c.setCellStyle(headerStyle);
        }
        for (int i = 0; i < taglie.size(); i++) {
            Cell c = header.createCell(i + 4);
            c.setCellValue(taglie.get(i));
            c.setCellStyle(headerStyle);
        }
        Cell totCell = header.createCell(4 + taglie.size());
        totCell.setCellValue("TOTALE");
        totCell.setCellStyle(headerStyle);

        Map<String, Map<String, Integer>> pivot = new LinkedHashMap<>();
        Map<String, String[]> chiavi = new LinkedHashMap<>();
        for (Riga r : righe) {
            String key = r.modello + "|" + r.tessuto + "|" + r.trattamento + "|" + r.colore;
            if (!pivot.containsKey(key)) {
                pivot.put(key, new HashMap<>());
                chiavi.put(key, new String[]{r.modello, r.tessuto, r.trattamento, r.colore});
            }
            pivot.get(key).put(r.taglia, r.giacenza);
        }

        int rowIdx = 2;
        int totaleMagazzino = 0;
        for (Map.Entry<String, Map<String, Integer>> entry : pivot.entrySet()) {
            String[] campi = chiavi.get(entry.getKey());
            Row row = sheet.createRow(rowIdx++);
            for (int j = 0; j < 4; j++) {
                Cell c = row.createCell(j);
                c.setCellValue(campi[j]);
                c.setCellStyle(dataStyle);
            }

            Map<String, Integer> tagliaGiac = entry.getValue();
            int totRiga = 0;
            for (int i = 0; i < taglie.size(); i++) {
                int giac = tagliaGiac.getOrDefault(taglie.get(i), 0);
                Cell c = row.createCell(i + 4);
                c.setCellValue(giac);
                c.setCellStyle(dataStyle);
                totRiga += giac;
            }

            Cell cTot = row.createCell(4 + taglie.size());
            cTot.setCellValue(totRiga);
            cTot.setCellStyle(dataStyle);
            totaleMagazzino += totRiga;
        }

        magazzinoRow.createCell(colTotMagazzino + 1).setCellValue(totaleMagazzino);
        magazzinoRow.getCell(colTotMagazzino + 1).setCellStyle(headerStyle);

        for (int i = 0; i <= colTotMagazzino + 1; i++) sheet.autoSizeColumn(i);
    }

    private static String getCellString(Row row, Integer col) {
        if (col == null || col < 0) return "";
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf((long) cell.getNumericCellValue());
        return cell.toString().trim();
    }
}
