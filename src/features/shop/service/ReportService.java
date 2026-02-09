package features.shop.service;

import core.db.DbManager;
import core.net.FtpUtils;
import features.shop.ui.ShopAppFrame;

import javax.swing.*;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URI;
import java.net.URLEncoder;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ReportService {
    private final DbManager dbManager;
    private final Properties props;
    private final File configFile; // non usato qui, ma utile se in futuro vuoi salvare

    public ReportService(DbManager dbManager, Properties props, File configFile) {
        this.dbManager = dbManager;
        this.props = props;
        this.configFile = configFile;
    }

    /**
     * Sequenza:
     * 1) esportazione statistiche complete (tutto, non filtrato per data)
     * 2) invio file via FTP (se configurato)
     * 3) trasmissione venduto del giorno selezionato via mail
     */
    public void esportaStatisticheETrasmettiVenduto(JFrame parent, String outPath, ShopAppFrame gui) {
        if (outPath == null) outPath = "";
        if (!outPath.toLowerCase().endsWith(".xlsx")) outPath = outPath + ".xlsx";

        String anagPath = gui.getProps().getProperty("lastTabellaModelli", "");

        if (dbManager == null || outPath.isEmpty() || anagPath.isEmpty()) {
            JOptionPane.showMessageDialog(parent,
                    "Controlla le impostazioni di connessione Firebird, il file di output e il file di raffronto!",
                    "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        gui.setStatButtonEnabled(false);
        gui.setProgressIndeterminate(false);
        gui.setProgressBarVisible(true);
        gui.setStatus("Esportazione statistiche complete in corso...");
        gui.setProgress(0);

        String finalOutPath = outPath;

        new SwingWorker<Void, Integer>() {
            private Exception exportException = null;
            private Exception ftpException = null;
            private boolean ftpSuccess = false;

            @Override
            protected Void doInBackground() {
                try {
                    runExportStatisticheComplete(finalOutPath, anagPath, this::publish, gui);
                } catch (Exception ex) {
                    exportException = ex;
                    return null;
                }

                // FTP
                String ftpHost = props.getProperty("ftpHost", "");
                String ftpUser = props.getProperty("ftpUser", "");
                String ftpPass = props.getProperty("ftpPassword", "");
                String ftpDir  = props.getProperty("ftpDir", "");
                boolean ftpEnabled = !ftpHost.isEmpty() && !ftpUser.isEmpty() && !ftpPass.isEmpty();

                if (ftpEnabled) {
                    String remoteFileName = new File(finalOutPath).getName();
                    try {
                        gui.appendLog("[LOG] Invio file via FTP a: " + ftpHost + (ftpDir.isEmpty() ? "" : ("/" + ftpDir)));
                        FtpUtils.uploadFileToFtp(finalOutPath, remoteFileName, ftpHost, ftpUser, ftpPass, ftpDir);
                        gui.appendLog("[LOG] Upload FTP completato.");
                        ftpSuccess = true;
                    } catch (Exception ftpEx) {
                        ftpException = ftpEx;
                        gui.appendLog("[ERR] Errore FTP: " + ftpEx.getMessage());
                    }
                } else {
                    gui.appendLog("[LOG] FTP non configurato. Nessun upload eseguito.");
                }

                return null;
            }

            @Override
            protected void process(java.util.List<Integer> chunks) {
                int last = chunks.get(chunks.size() - 1);
                gui.setProgressIndeterminate(false);
                gui.setProgress(last);
                gui.setStatus("Avanzamento esportazione statistiche complete: " + last + "%");
            }

            @Override
            protected void done() {
                gui.setProgressBarVisible(false);
                gui.setStatButtonEnabled(true);

                if (exportException != null) {
                    gui.setStatus("Errore esportazione statistiche complete!");
                    gui.appendLog("[ERRORE] (Statistiche Complete) " + exportException);
                    JOptionPane.showMessageDialog(parent, "Errore: " + exportException.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                if (ftpException != null) {
                    gui.setStatus("Errore invio file via FTP!");
                    JOptionPane.showMessageDialog(parent,
                            "Errore durante l'invio del file via FTP:\n" + ftpException.getMessage(),
                            "Errore FTP", JOptionPane.ERROR_MESSAGE);
                } else if (ftpSuccess) {
                    gui.setStatus("Esportazione + invio FTP completati!");
                    JOptionPane.showMessageDialog(parent,
                            "Esportazione statistiche complete completata e upload FTP riuscito!",
                            "Esito", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    gui.setStatus("Esportazione completata (FTP non configurato)");
                    JOptionPane.showMessageDialog(parent,
                            "Esportazione statistiche complete completata!",
                            "Esito", JOptionPane.INFORMATION_MESSAGE);
                }

                trasmettiVendutoGiornaliero(parent, gui);
            }
        }.execute();
    }

    private void runExportStatisticheComplete(String outPath, String anagPath, java.util.function.Consumer<Integer> progressCallback, ShopAppFrame gui) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            int totalSteps = 3;
            int currentStep = 0;

            esportaStatisticheVendutoSuFoglio(wb, "Stat Vendita", gui, val -> progressCallback.accept(val / totalSteps));
            currentStep++;
            progressCallback.accept(currentStep * 100 / totalSteps);

            esportaScaricoAggregatoSuFoglio(wb, "Stat Capi", anagPath, gui, val -> progressCallback.accept(33 + val / totalSteps));
            currentStep++;
            progressCallback.accept(currentStep * 100 / totalSteps);

            esportaDettaglioScarico(wb, "Dettaglio Scarico", anagPath, gui, val -> progressCallback.accept(66 + val / totalSteps));

            progressCallback.accept(100);

            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                wb.write(fos);
            }
        }

        progressCallback.accept(100);
        gui.appendLog("[LOG] Esportazione statistiche complete completata!");
    }

    private void esportaStatisticheVendutoSuFoglio(XSSFWorkbook wb, String sheetName, ShopAppFrame gui, java.util.function.IntConsumer progressCallback) throws Exception {
        try (Connection conn = dbManager.getConnection()) {
            String countSql = "SELECT COUNT(*) FROM \"TDocTestate\"";
            int totalRows = 1;
            try (Statement st = conn.createStatement(); ResultSet rs = st.executeQuery(countSql)) {
                if (rs.next()) totalRows = rs.getInt(1);
            }

            String sql = """
                SELECT EXTRACT(DAY FROM "Data") as "Giorno",
                       EXTRACT(MONTH FROM "Data") as "Mese",
                       EXTRACT(YEAR FROM "Data") as "Anno",
                       SUM("TotNetto") as "Netto",
                       SUM("TotIva") as "IVA",
                       SUM("TotDoc") as "Totale"
                FROM (
                    SELECT "Data", "TotNetto", "TotIva", "TotDoc"
                    FROM "TDocTestate"
                )
                GROUP BY "Data"
                ORDER BY "Anno", "Mese", "Giorno"
            """;

            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet rs = ps.executeQuery()) {

                XSSFSheet sheet = wb.createSheet(sheetName);
                String[] columns = {"Giorno-Mese-Anno", "Netto", "IVA", "Totale"};
                Row header = sheet.createRow(0);
                for (int i = 0; i < columns.length; i++) header.createCell(i).setCellValue(columns[i]);

                XSSFCellStyle numberStyle = wb.createCellStyle();
                XSSFDataFormat dfmt = wb.createDataFormat();
                numberStyle.setDataFormat(dfmt.getFormat("#,##0.00"));

                int rowNum = 1;
                int processed = 0;

                while (rs.next()) {
                    String giorno = rs.getString("Giorno");
                    String mese = rs.getString("Mese");
                    String anno = rs.getString("Anno");
                    String data = (giorno.length() == 1 ? "0" : "") + giorno + "-" + (mese.length() == 1 ? "0" : "") + mese + "-" + anno;

                    double netto = rs.getDouble("Netto");
                    double iva = rs.getDouble("IVA");
                    double totale = rs.getDouble("Totale");

                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(data);

                    Cell c1 = row.createCell(1); c1.setCellValue(netto); c1.setCellStyle(numberStyle);
                    Cell c2 = row.createCell(2); c2.setCellValue(iva); c2.setCellStyle(numberStyle);
                    Cell c3 = row.createCell(3); c3.setCellValue(totale); c3.setCellStyle(numberStyle);

                    processed++;
                    if (processed % 10 == 0 || processed == totalRows) {
                        progressCallback.accept((int) (100.0 * processed / totalRows));
                    }
                }

                for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
                progressCallback.accept(100);
            }
        }
    }

    private void esportaScaricoAggregatoSuFoglio(XSSFWorkbook wb, String sheetName, String anagPath, ShopAppFrame gui, java.util.function.IntConsumer progressCallback) throws Exception {
        Map<String, String[]> barcodeToGenereTipologia = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(anagPath);
             Workbook anagWb = new XSSFWorkbook(fis)) {

            Sheet sheet = anagWb.getSheetAt(0);
            Map<String, Integer> colIndex = new HashMap<>();
            Row header = sheet.getRow(0);
            for (Cell cell : header) {
                String v = cell.getStringCellValue().trim().toUpperCase();
                colIndex.put(v, cell.getColumnIndex());
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String barcode = getCellString(row, colIndex.getOrDefault("BARCODE", -1));
                String genere = getCellString(row, colIndex.getOrDefault("GENERE", -1));
                String tipologia = getCellString(row, colIndex.getOrDefault("TIPOLOGIA", -1));
                if (!barcode.isEmpty()) barcodeToGenereTipologia.put(barcode, new String[]{genere, tipologia});
            }
        }

        String countSql = """
            SELECT COUNT(*) FROM "TMovMagazz"
            WHERE "QtaScaricata" IS NOT NULL AND "QtaScaricata" > 0
            AND "PrezzoNetto" IS NOT NULL AND "PrezzoNetto" > 0
        """;

        int totalRows = 1;
        try (Connection conn = dbManager.getConnection();
             Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery(countSql)) {
            if (rs.next()) totalRows = rs.getInt(1);
        }

        String sql = """
            SELECT
                "TMovMagazz"."Data" as DATA,
                "TArticoliTaglieColori"."CodBarre" AS EAN,
                "TMovMagazz"."QtaScaricata" AS Scaricato
            FROM
                "TMovMagazz"
            LEFT JOIN "TArticoliTaglieColori"
                ON "TArticoliTaglieColori"."IDArticolo" = "TMovMagazz"."IDArticolo"
                AND "TArticoliTaglieColori"."Taglia" = SUBSTRING("TMovMagazz"."Lotto" FROM 1 FOR POSITION('/', "TMovMagazz"."Lotto") - 1)
                AND "TArticoliTaglieColori"."Colore" = SUBSTRING("TMovMagazz"."Lotto" FROM POSITION('/', "TMovMagazz"."Lotto") + 1)
            WHERE
                "TMovMagazz"."QtaScaricata" IS NOT NULL
                AND "TMovMagazz"."QtaScaricata" > 0
                AND "TMovMagazz"."PrezzoNetto" IS NOT NULL
                AND "TMovMagazz"."PrezzoNetto" > 0
            ORDER BY
                "TMovMagazz"."Data"
        """;

        Map<java.util.Date, int[]> aggregati = new TreeMap<>();

        try (Connection conn = dbManager.getConnection();
             PreparedStatement ps = conn.prepareStatement(sql);
             ResultSet rs = ps.executeQuery()) {

            int processed = 0;
            while (rs.next()) {
                java.util.Date dataSql = rs.getDate("DATA");
                String ean = rs.getString("EAN");
                int scaricato = rs.getInt("Scaricato");

                String[] gruppo = barcodeToGenereTipologia.get(ean);
                String genere = gruppo != null ? (gruppo[0] == null ? "" : gruppo[0].trim().toUpperCase()) : "";
                String tipologia = gruppo != null ? (gruppo[1] == null ? "" : gruppo[1].trim().toUpperCase()) : "";

                int idx;
                if (genere.equals("UOMO") && tipologia.equals("PANTALONE")) idx = 0;
                else if (genere.equals("DONNA") && tipologia.equals("PANTALONE")) idx = 1;
                else idx = 2;

                int[] arr = aggregati.computeIfAbsent(dataSql, k -> new int[3]);
                arr[idx] += scaricato;

                processed++;
                if (processed % 100 == 0 || processed == totalRows) {
                    progressCallback.accept((int) (100.0 * processed / totalRows));
                }
            }
        }

        XSSFSheet sheet = wb.createSheet(sheetName);
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("DATA");
        header.createCell(1).setCellValue("PANT. UOMO");
        header.createCell(2).setCellValue("PANT. DONNA");
        header.createCell(3).setCellValue("ALTRO");

        int rowIdx = 1;
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
        for (Map.Entry<java.util.Date, int[]> entry : aggregati.entrySet()) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(sdf.format(entry.getKey()));
            row.createCell(1).setCellValue(entry.getValue()[0]);
            row.createCell(2).setCellValue(entry.getValue()[1]);
            row.createCell(3).setCellValue(entry.getValue()[2]);
        }

        for (int i = 0; i < 4; i++) sheet.autoSizeColumn(i);
        progressCallback.accept(100);
    }

    private void esportaDettaglioScarico(XSSFWorkbook wb, String sheetName, String anagPath, ShopAppFrame gui, java.util.function.IntConsumer progressCallback) throws Exception {
        Map<String, String[]> barcodeToInfo = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(anagPath);
             Workbook anagWb = new XSSFWorkbook(fis)) {

            Sheet sheet = anagWb.getSheetAt(0);
            Map<String, Integer> colIndex = new HashMap<>();
            Row header = sheet.getRow(0);
            for (Cell cell : header) {
                String v = cell.getStringCellValue().trim().toUpperCase();
                colIndex.put(v, cell.getColumnIndex());
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String barcode = getCellString(row, colIndex.getOrDefault("BARCODE", -1));
                String modello = getCellString(row, colIndex.getOrDefault("MODELLO", -1));
                String tessuto = getCellString(row, colIndex.getOrDefault("TESSUTO", -1));
                String trattamento = getCellString(row, colIndex.getOrDefault("TRATTAMENTO", -1));
                String colore = getCellString(row, colIndex.getOrDefault("COLORE", -1));
                String taglia = getCellString(row, colIndex.getOrDefault("TAGLIA", -1));

                if (!barcode.isEmpty()) barcodeToInfo.put(barcode, new String[]{modello, tessuto, trattamento, colore, taglia});
            }
        }

        String sql = """
            SELECT
                "TMovMagazz"."Data" AS DATA,
                "TArticoliTaglieColori"."CodBarre" AS EAN,
                "TMovMagazz"."QtaScaricata" AS Scaricato,
                "TMovMagazz"."PrezzoNetto" AS PrezzoNetto
            FROM
                "TMovMagazz"
            LEFT JOIN "TArticoliTaglieColori"
                ON "TArticoliTaglieColori"."IDArticolo" = "TMovMagazz"."IDArticolo"
                AND "TArticoliTaglieColori"."Taglia" = SUBSTRING("TMovMagazz"."Lotto" FROM 1 FOR POSITION('/', "TMovMagazz"."Lotto") - 1)
                AND "TArticoliTaglieColori"."Colore" = SUBSTRING("TMovMagazz"."Lotto" FROM POSITION('/', "TMovMagazz"."Lotto") + 1)
            WHERE
                "TMovMagazz"."QtaScaricata" IS NOT NULL
                AND "TMovMagazz"."QtaScaricata" > 0
                AND "TMovMagazz"."PrezzoNetto" IS NOT NULL
                AND "TMovMagazz"."PrezzoNetto" > 0
        """;

        XSSFSheet sheet = wb.createSheet(sheetName);
        Row header = sheet.createRow(0);
        String[] columns = {"DATA","EAN","MODELLO","TESSUTO","TRATTAMENTO","COLORE","TAGLIA","SCARICATO","PREZZO"};
        for (int i = 0; i < columns.length; i++) header.createCell(i).setCellValue(columns[i]);

        int rowIdx = 1;
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

        try (Connection conn = dbManager.getConnection();
             PreparedStatement ps = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
             ResultSet rs = ps.executeQuery()) {

            rs.last(); int total = rs.getRow(); rs.beforeFirst();
            int processed = 0;

            while (rs.next()) {
                Row row = sheet.createRow(rowIdx++);

                java.util.Date dataSql = rs.getDate("DATA");
                String ean = rs.getString("EAN");
                int scaricato = rs.getInt("Scaricato");
                double prezzoNetto = rs.getDouble("PrezzoNetto");
                double prezzo = Math.round(prezzoNetto * 1.22 * 100.0) / 100.0;

                String[] info = barcodeToInfo.getOrDefault(ean, new String[]{"","","","",""});

                row.createCell(0).setCellValue(sdf.format(dataSql));
                row.createCell(1).setCellValue(ean != null ? ean : "");
                row.createCell(2).setCellValue(info[0]);
                row.createCell(3).setCellValue(info[1]);
                row.createCell(4).setCellValue(info[2]);
                row.createCell(5).setCellValue(info[3]);
                row.createCell(6).setCellValue(info[4]);
                row.createCell(7).setCellValue(scaricato);
                row.createCell(8).setCellValue(prezzo);

                processed++;
                if (total > 0 && (processed % 100 == 0 || processed == total)) {
                    progressCallback.accept((int) (100.0 * processed / total));
                }
            }
        }

        for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
        progressCallback.accept(100);
    }

    public void trasmettiVendutoGiornaliero(JFrame parent, ShopAppFrame gui) {
        if (dbManager == null) {
            JOptionPane.showMessageDialog(parent, "Controlla le impostazioni di connessione Firebird!", "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        java.util.Date selectedDate = (java.util.Date) gui.getDatePicker().getModel().getValue();
        if (selectedDate == null) {
            JOptionPane.showMessageDialog(parent, "Seleziona una data!", "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        LocalDate sel = selectedDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        LocalDate today = LocalDate.now();
        if (sel.isAfter(today)) {
            JOptionPane.showMessageDialog(parent, "Non puoi selezionare una data futura!", "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        java.sql.Date sqlDate = java.sql.Date.valueOf(sel);
        String dateStr = String.format("%02d-%02d-%04d", sel.getDayOfMonth(), sel.getMonthValue(), sel.getYear());

        ProgressMonitor pm = new ProgressMonitor(parent, "Elaborazione venduto giornaliero...", "Avvio...", 0, 100);
        pm.setMillisToDecideToPopup(50);

        SwingWorker<Void, Integer> worker = new SwingWorker<>() {
            StringBuilder sb = new StringBuilder();

            @Override
            protected Void doInBackground() {
                publish(0);
                sb.append("Venduto per data ").append(dateStr).append(":\n\n");

                double totaleGenerale = 0.0;

                try (Connection conn = dbManager.getConnection()) {
                    String sql = """
                        SELECT
                          CASE
                            WHEN "DescDoc" LIKE 'Vend. banco%' THEN 'Vendita da Banco'
                            WHEN "DescDoc" LIKE 'Fattura P.F.%' THEN 'Fatture Pro Forma'
                          END AS Gruppo,
                          "Pagamento",
                          SUM("TotDoc") AS TotaleVenduto
                        FROM "TDocTestate"
                        WHERE "Data" = ?
                          AND ("DescDoc" LIKE 'Vend. banco%' OR "DescDoc" LIKE 'Fattura P.F.%')
                        GROUP BY Gruppo, "Pagamento"
                        ORDER BY Gruppo, "Pagamento"
                    """;

                    try (PreparedStatement ps = conn.prepareStatement(sql)) {
                        ps.setDate(1, sqlDate);

                        try (ResultSet rs = ps.executeQuery()) {
                            Map<String, List<String>> righePerGruppo = new LinkedHashMap<>();
                            Map<String, Double> totaliPerGruppo = new LinkedHashMap<>();

                            DecimalFormatSymbols symbols = new DecimalFormatSymbols();
                            symbols.setDecimalSeparator(',');
                            DecimalFormat df = new DecimalFormat("#,##0.00", symbols);

                            while (rs.next()) {
                                String gruppo = rs.getString("Gruppo");
                                String pagamento = rs.getString("Pagamento");
                                double totale = rs.getDouble("TotaleVenduto");

                                righePerGruppo.computeIfAbsent(gruppo, k -> new ArrayList<>())
                                        .add("  " + (pagamento == null ? "N/D" : pagamento) + ": " + df.format(totale));
                                totaliPerGruppo.put(gruppo, totaliPerGruppo.getOrDefault(gruppo, 0.0) + totale);

                                totaleGenerale += totale;
                            }

                            if (righePerGruppo.isEmpty()) {
                                sb.append("Nessun dato trovato per questa data.");
                            } else {
                                for (Map.Entry<String, List<String>> entry : righePerGruppo.entrySet()) {
                                    sb.append(entry.getKey()).append("\n");
                                    for (String riga : entry.getValue()) sb.append(riga).append("\n");
                                    sb.append("  TOTALE ").append(entry.getKey().toUpperCase()).append(": ")
                                            .append(df.format(totaliPerGruppo.get(entry.getKey()))).append("\n\n");
                                }
                                sb.append("\nTOTALE GENERALE: ").append(df.format(totaleGenerale));
                            }
                        }
                    }
                } catch (Exception ex) {
                    gui.appendLog("[ERRORE] (Venduto Giornaliero) " + ex);
                    sb.append("ERRORE: ").append(ex.getMessage());
                }

                publish(40);

                try {
                    String anagPath = gui.getProps().getProperty("lastTabellaModelli", "");
                    Map<String, String[]> barcodeToGenereTipologia = new HashMap<>();

                    try (FileInputStream fis = new FileInputStream(anagPath);
                         Workbook anagWb = new XSSFWorkbook(fis)) {

                        Sheet sheet = anagWb.getSheetAt(0);
                        Map<String, Integer> colIndex = new HashMap<>();
                        Row header = sheet.getRow(0);
                        for (Cell cell : header) {
                            String v = cell.getStringCellValue().trim().toUpperCase();
                            colIndex.put(v, cell.getColumnIndex());
                        }

                        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                            Row row = sheet.getRow(r);
                            if (row == null) continue;

                            String barcode = getCellString(row, colIndex.getOrDefault("BARCODE", -1));
                            String genere = getCellString(row, colIndex.getOrDefault("GENERE", -1));
                            String tipologia = getCellString(row, colIndex.getOrDefault("TIPOLOGIA", -1));
                            if (!barcode.isEmpty()) barcodeToGenereTipologia.put(barcode, new String[]{genere, tipologia});
                        }
                    }

                    publish(60);

                    int[] capi = new int[3];

                    try (Connection conn = dbManager.getConnection()) {
                        String sql = """
                            SELECT
                                "TArticoliTaglieColori"."CodBarre" AS EAN,
                                "TMovMagazz"."QtaScaricata" AS Scaricato
                            FROM
                                "TMovMagazz"
                            LEFT JOIN "TArticoliTaglieColori"
                                ON "TArticoliTaglieColori"."IDArticolo" = "TMovMagazz"."IDArticolo"
                                AND "TArticoliTaglieColori"."Taglia" = SUBSTRING("TMovMagazz"."Lotto" FROM 1 FOR POSITION('/', "TMovMagazz"."Lotto") - 1)
                                AND "TArticoliTaglieColori"."Colore" = SUBSTRING("TMovMagazz"."Lotto" FROM POSITION('/', "TMovMagazz"."Lotto") + 1)
                            WHERE
                                "TMovMagazz"."QtaScaricata" IS NOT NULL
                                AND "TMovMagazz"."QtaScaricata" > 0
                                AND "TMovMagazz"."PrezzoNetto" IS NOT NULL
                                AND "TMovMagazz"."PrezzoNetto" > 0
                                AND "TMovMagazz"."Data" = ?
                        """;

                        try (PreparedStatement ps = conn.prepareStatement(sql)) {
                            ps.setDate(1, sqlDate);
                            try (ResultSet rs = ps.executeQuery()) {
                                while (rs.next()) {
                                    String ean = rs.getString("EAN");
                                    int scaricato = rs.getInt("Scaricato");

                                    String[] gruppo = barcodeToGenereTipologia.get(ean);
                                    String genere = gruppo != null ? (gruppo[0] == null ? "" : gruppo[0].trim().toUpperCase()) : "";
                                    String tipologia = gruppo != null ? (gruppo[1] == null ? "" : gruppo[1].trim().toUpperCase()) : "";

                                    int idx;
                                    if (genere.equals("UOMO") && tipologia.equals("PANTALONE")) idx = 0;
                                    else if (genere.equals("DONNA") && tipologia.equals("PANTALONE")) idx = 1;
                                    else idx = 2;

                                    capi[idx] += scaricato;
                                }
                            }
                        }
                    }

                    sb.append("\nCAPI UOMO: ").append(capi[0]);
                    sb.append("\nCAPI DONNA: ").append(capi[1]);
                    sb.append("\nALTRO: ").append(capi[2]).append("\n");
                } catch (Exception ex) {
                    gui.appendLog("[ERRORE] (Capi Venduto Giornaliero) " + ex);
                    sb.append("\n[ERRORE lettura capi per la data selezionata]");
                }

                publish(100);
                return null;
            }

            @Override
            protected void process(List<Integer> chunks) {
                int value = chunks.get(chunks.size() - 1);
                pm.setProgress(value);
                pm.setNote("Avanzamento: " + value + "%");
            }

            @Override
            protected void done() {
                pm.close();
                String[] destinatari = gui.chiediDestinatariVenduto();
                if (destinatari == null) return;

                String to = destinatari[0];
                String cc = destinatari[1];
                String mittente = destinatari[2];

                try {
                    String subject = "Venduto giornaliero - " + mittente + " - " + dateStr;
                    String body = sb.toString();

                    String encodedSubject = URLEncoder.encode(subject, "UTF-8").replace("+", "%20");
                    String encodedBody = URLEncoder.encode(body, "UTF-8").replace("+", "%20");
                    String encodedCc = URLEncoder.encode(cc, "UTF-8").replace("+", "%20");

                    String mailto = String.format("mailto:%s?cc=%s&subject=%s&body=%s", to, encodedCc, encodedSubject, encodedBody);
                    Desktop.getDesktop().mail(new URI(mailto));
                } catch (Exception ex) {
                    gui.appendLog("[ERRORE] (Mail) " + ex);
                    JOptionPane.showMessageDialog(parent,
                            "Impossibile creare la mail automatica:\n\n" + ex.getMessage(),
                            "Errore", JOptionPane.ERROR_MESSAGE);
                }
            }
        };

        worker.execute();
    }

    private static String getCellString(Row row, Integer col) {
        if (col == null || col < 0) return "";
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf((long) cell.getNumericCellValue());
        return cell.toString().trim();
    }
}
