package features.shop.ui;

import core.config.ConfigUtils;
import core.db.DbManager;
import core.db.FirebirdSettings;
import core.ui.LogWindow;
import features.shop.service.MagazzinoService;
import features.shop.service.ReportService;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.Properties;

import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.UtilDateModel;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.DateComponentFormatter;

/**
 * Porting della vecchia gui.AppFrame (Shop-App) dentro la nuova struttura v2.
 */
public class ShopAppFrame {
    private JFrame frame;

    private JButton startButton, statButton;
    private JDatePickerImpl datePicker;
    private JProgressBar progressBar;
    private JLabel statusLabel, connLabel;

    private final File configFile = new File("config.properties");
    private final Properties props = new Properties();

    private final FirebirdSettings fbSettings = new FirebirdSettings();
    private DbManager dbManager;

    private final LogWindow logWindow = new LogWindow();
    private final StringBuilder sessionLog = new StringBuilder();

    private MagazzinoService magazzinoService;
    private ReportService reportService;

    public void showUI() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ignored) {}

        loadConfig();
        createServices();

        frame = new JFrame("Tramarossa Statistiche Negozi");

        // Icona (se presente)
        try {
            ImageIcon favicon = new ImageIcon("img" + File.separator + "favico.jpg");
            frame.setIconImage(favicon.getImage());
        } catch (Exception ignored) {}

        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setSize(950, 600);
        frame.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();

        // Spacer top
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 3; gbc.weighty = 0;
        JLabel topSpacer = new JLabel();
        topSpacer.setPreferredSize(new Dimension(1, 100));
        frame.add(topSpacer, gbc);

        // Logo
        JLabel logoLabel = new JLabel();
        logoLabel.setHorizontalAlignment(SwingConstants.CENTER);
        try {
            ImageIcon logoIcon = new ImageIcon("img" + File.separator + "tramarossa.png");
            Image logoImg = logoIcon.getImage().getScaledInstance(750, 90, Image.SCALE_SMOOTH);
            logoLabel.setIcon(new ImageIcon(logoImg));
        } catch (Exception ex) {
            logoLabel.setText("Tramarossa");
        }
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 3;
        gbc.insets = new Insets(0, 0, 5, 0);
        gbc.anchor = GridBagConstraints.NORTH;
        frame.add(logoLabel, gbc);

        // Conn label
        connLabel = new JLabel();
        aggiornaConnLabel();
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 3;
        gbc.insets = new Insets(0, 10, 5, 10);
        gbc.anchor = GridBagConstraints.CENTER;
        frame.add(connLabel, gbc);

        // Spacer between conn and buttons
        gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 3; gbc.weighty = 0.20;
        frame.add(new JLabel(""), gbc);

        // Buttons
        startButton = new JButton("Esporta Magazzino");
        gbc.gridx = 1; gbc.gridy = 4; gbc.gridwidth = 1; gbc.insets = new Insets(0,0,0,0); gbc.weighty = 0;
        frame.add(startButton, gbc);

        statButton = new JButton("Trasmetti Vendite Giornaliere");
        gbc.gridx = 2; gbc.gridy = 4;
        frame.add(statButton, gbc);

        // Date picker
        gbc.gridx = 0; gbc.gridy = 5; gbc.anchor = GridBagConstraints.EAST; gbc.insets = new Insets(10,10,5,5);
        frame.add(new JLabel("Data da trasmettere:"), gbc);

        UtilDateModel model = new UtilDateModel();
        LocalDate oggi = LocalDate.now();
        model.setDate(oggi.getYear(), oggi.getMonthValue()-1, oggi.getDayOfMonth());
        model.setSelected(true);
        Properties p = new Properties();
        JDatePanelImpl datePanel = new JDatePanelImpl(model, p);
        datePicker = new JDatePickerImpl(datePanel, new DateComponentFormatter());
        gbc.gridx = 1; gbc.anchor = GridBagConstraints.WEST; gbc.insets = new Insets(10,0,5,0);
        frame.add(datePicker, gbc);

        // Progress + status
        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);
        progressBar.setVisible(false);
        gbc.gridx = 1; gbc.gridy = 6; gbc.gridwidth = 2; gbc.insets = new Insets(10,0,0,0);
        frame.add(progressBar, gbc);

        statusLabel = new JLabel(" ");
        gbc.gridx = 0; gbc.gridy = 7; gbc.gridwidth = 3; gbc.insets = new Insets(5,0,0,0);
        frame.add(statusLabel, gbc);

        gbc.gridx = 0; gbc.gridy = 8; gbc.gridwidth = 3; gbc.weighty = 1.0;
        frame.add(new JLabel(""), gbc);

        // Menu
        JMenuBar menuBar = new JMenuBar();
        JMenu fileMenu = new JMenu("File");
        JMenuItem settingsItem = new JMenuItem("Impostazioni...");
        JMenuItem logItem = new JMenuItem("Log...");
        JMenuItem exitItem = new JMenuItem("Chiudi");
        fileMenu.add(settingsItem);
        fileMenu.add(logItem);
        fileMenu.addSeparator();
        fileMenu.add(exitItem);
        menuBar.add(fileMenu);
        frame.setJMenuBar(menuBar);

        // Listeners
        startButton.addActionListener(e -> magazzinoService.exportMagazzino(
                frame,
                props.getProperty("lastExportMagazzino", ""),
                props.getProperty("lastTabellaModelli", ""),
                this
        ));
        statButton.addActionListener(e -> reportService.esportaStatisticheETrasmettiVenduto(
                frame,
                props.getProperty("lastExportStatistiche", ""),
                this
        ));

        settingsItem.addActionListener(e -> mostraDialogImpostazioni());
        logItem.addActionListener(e -> logWindow.setVisible(true));
        exitItem.addActionListener(e -> frame.dispose());

        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    private void createServices() {
        dbManager = new DbManager(fbSettings);
        magazzinoService = new MagazzinoService(dbManager, props, configFile);
        reportService = new ReportService(dbManager, props, configFile);
    }

    private void loadConfig() {
        try {
            props.putAll(ConfigUtils.load(configFile));
        } catch (Exception ignored) {}

        fbSettings.setDbPath(props.getProperty("fb_dbpath", ""));
        fbSettings.setUser(props.getProperty("fb_user", "SYSDBA"));
        fbSettings.setPassword(props.getProperty("fb_password", "masterkey"));
        fbSettings.setJdbcUrlOverride(props.getProperty("fb_jdbcUrl", ""));
    }

    private void saveConfig() {
        try {
            ConfigUtils.save(configFile, props, "Tramarossa ToolSuite config");
        } catch (Exception ex) {
            appendLog("[ERR] Impossibile salvare config: " + ex.getMessage());
        }
    }

    private void mostraDialogImpostazioni() {
        JTextField dbPathField = new JTextField(fbSettings.getDbPath(), 25);
        JTextField jdbcOverrideField = new JTextField(fbSettings.getJdbcUrlOverride() == null ? "" : fbSettings.getJdbcUrlOverride(), 25);

        JTextField userField = new JTextField(fbSettings.getUser(), 15);
        JPasswordField passField = new JPasswordField(fbSettings.getPassword(), 15);

        JTextField outFieldSettings = new JTextField(props.getProperty("lastExportMagazzino", ""), 25);
        JTextField statFieldSettings = new JTextField(props.getProperty("lastExportStatistiche", ""), 25);
        JTextField anagFieldSettings = new JTextField(props.getProperty("lastTabellaModelli", ""), 25);

        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4,4,4,4);

        int y = 0;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("DB Path (.eft):"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(dbPathField, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("JDBC URL override (opz):"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(jdbcOverrideField, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("Utente:"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(userField, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("Password:"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(passField, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("Output Excel Magazzino:"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(outFieldSettings, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("Output Statistiche:"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(statFieldSettings, gbc);

        y++;
        gbc.gridy=y; gbc.gridx=0; gbc.anchor=GridBagConstraints.EAST;
        panel.add(new JLabel("File Raffronto (BARCODE):"), gbc);
        gbc.gridx=1; gbc.anchor=GridBagConstraints.WEST;
        panel.add(anagFieldSettings, gbc);

        int res = JOptionPane.showConfirmDialog(
                frame, panel, "Impostazioni",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE
        );
        if (res == JOptionPane.OK_OPTION) {
            fbSettings.setDbPath(dbPathField.getText().trim());
            fbSettings.setJdbcUrlOverride(jdbcOverrideField.getText().trim());
            fbSettings.setUser(userField.getText().trim());
            fbSettings.setPassword(new String(passField.getPassword()));

            props.setProperty("fb_dbpath", fbSettings.getDbPath() == null ? "" : fbSettings.getDbPath());
            props.setProperty("fb_jdbcUrl", fbSettings.getJdbcUrlOverride() == null ? "" : fbSettings.getJdbcUrlOverride());
            props.setProperty("fb_user", fbSettings.getUser() == null ? "" : fbSettings.getUser());
            props.setProperty("fb_password", fbSettings.getPassword() == null ? "" : fbSettings.getPassword());

            props.setProperty("lastExportMagazzino", outFieldSettings.getText().trim());
            props.setProperty("lastExportStatistiche", statFieldSettings.getText().trim());
            props.setProperty("lastTabellaModelli", anagFieldSettings.getText().trim());

            saveConfig();
            aggiornaConnLabel();
            appendLog("[LOG] Impostazioni aggiornate.");
        }
    }

    private void aggiornaConnLabel() {
        String mode = (fbSettings.getJdbcUrlOverride() != null && !fbSettings.getJdbcUrlOverride().trim().isEmpty())
                ? "JDBC URL"
                : "EMBEDDED";

        String db = (mode.equals("JDBC URL")) ? fbSettings.getJdbcUrlOverride() : fbSettings.getDbPath();

        String s = String.format(
                "<html><b>Connessione Firebird:</b> <font color='blue'>%s</font><br/>DB: <font color='blue'>%s</font> - Utente: <font color='blue'>%s</font></html>",
                mode, db, fbSettings.getUser()
        );
        if (connLabel != null) connLabel.setText(s);
    }

    // --- API usata dai service (come prima) ---
    public JDatePickerImpl getDatePicker() { return datePicker; }
    public void setProgressBarVisible(boolean vis) { progressBar.setVisible(vis); }
    public void setProgressIndeterminate(boolean ind) { progressBar.setIndeterminate(ind); }
    public void setProgress(int value) { progressBar.setValue(value); }
    public void setStatus(String txt) { statusLabel.setText(txt); }
    public void setStartButtonEnabled(boolean enabled) { startButton.setEnabled(enabled); }
    public void setStatButtonEnabled(boolean enabled) { statButton.setEnabled(enabled); }

    public void appendLog(String msg) {
        String line = "[" + new SimpleDateFormat("HH:mm:ss").format(new Date()) + "] " + msg;
        sessionLog.append(line).append("\n");
        logWindow.append(line);
        System.out.println(line);
    }

    public JFrame getFrame() { return frame; }

    public String[] chiediDestinatariVenduto() {
        String lastTo = props.getProperty("lastVendutoTo", "roberto.chemello@tramarossa.it");
        String lastCc = props.getProperty("lastVendutoCc", "paolo.chemello@tramarossa.it");
        String lastMittente = props.getProperty("lastVendutoMittente", "Tramarossa Treviso");

        JTextField toField = new JTextField(lastTo, 30);
        JTextField ccField = new JTextField(lastCc, 30);
        JTextField mittenteField = new JTextField(lastMittente, 30);

        JPanel panel = new JPanel(new GridLayout(0, 1));
        panel.add(new JLabel("A:"));
        panel.add(toField);
        panel.add(new JLabel("CC:"));
        panel.add(ccField);
        panel.add(new JLabel("Mittente:"));
        panel.add(mittenteField);

        int result = JOptionPane.showConfirmDialog(
                frame, panel, "Destinatari email venduto giornaliero",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE
        );

        if (result == JOptionPane.OK_OPTION) {
            String to = toField.getText().trim();
            String cc = ccField.getText().trim();
            String mittente = mittenteField.getText().trim();

            props.setProperty("lastVendutoTo", to);
            props.setProperty("lastVendutoCc", cc);
            props.setProperty("lastVendutoMittente", mittente);
            saveConfig();

            return new String[]{to, cc, mittente};
        }
        return null;
    }

    public FirebirdSettings getFbSettings() { return fbSettings; }
    public Properties getProps() { return props; }
}
