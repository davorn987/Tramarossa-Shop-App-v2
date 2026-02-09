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
import java.util.function.IntConsumer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

public class MagazzinoService {
    private final DbManager dbManager;
    private final Properties props;
    private final java.io.File configFile; // (non usata qui)

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
                    runElaborazione(outPath, anagPath, p -> publish(p), gui); // <-- QUI
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

    public void runElaborazione(String outPath, String anagPath, IntConsumer progressCallback, ShopAppFrame gui) throws Exception {
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

        // ... resto IDENTICO a prima ...
        // (lascia invariato)

        // NB: dove avevi progressCallback.accept(100) rimane uguale
    }

    private static String getCellString(Row row, Integer col) {
        if (col == null || col < 0) return "";
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf((long) cell.getNumericCellValue());
        return cell.toString().trim();
    }
}
