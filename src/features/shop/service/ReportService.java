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
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;
import java.util.function.IntConsumer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ReportService {
    private final DbManager dbManager;
    private final Properties props;
    private final File configFile;

    public ReportService(DbManager dbManager, Properties props, File configFile) {
        this.dbManager = dbManager;
        this.props = props;
        this.configFile = configFile;
    }

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

            @Override
            protected Void doInBackground() {
                try {
                    runExportStatisticheComplete(finalOutPath, anagPath, p -> publish(p), gui); // <-- QUI
                } catch (Exception ex) {
                    exportException = ex;
                }
                return null;
            }

            @Override
            protected void process(List<Integer> chunks) {
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

                // qui rimane la tua parte FTP + venduto (come già inviata)
            }
        }.execute();
    }

    private void runExportStatisticheComplete(String outPath, String anagPath, IntConsumer progressCallback, ShopAppFrame gui) throws Exception {
        // corpo invariato, ma usa progressCallback.accept(x)
    }

    // ... resto invariato ...
}
