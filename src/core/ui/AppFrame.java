package core.ui;

import javax.swing.*;
import java.awt.*;

public class AppFrame extends JFrame {
    private final LogWindow logWindow = new LogWindow();

    public AppFrame() {
        super("Tramarossa ToolSuite");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setSize(1000, 650);
        setLocationRelativeTo(null);

        setJMenuBar(buildMenu());

        JPanel center = new JPanel(new BorderLayout());
        center.add(new JLabel("Home - scegli una funzione dal menu", SwingConstants.CENTER), BorderLayout.CENTER);
        setContentPane(center);
    }

    private JMenuBar buildMenu() {
        JMenuBar mb = new JMenuBar();

        JMenu mFile = new JMenu("File");
        JMenuItem miExit = new JMenuItem("Esci");
        miExit.addActionListener(e -> dispose());
        mFile.add(miExit);

        JMenu mView = new JMenu("Vista");
        JMenuItem miLog = new JMenuItem("Log...");
        miLog.addActionListener(e -> logWindow.setVisible(true));
        mView.add(miLog);

        JMenu mFeatures = new JMenu("Funzioni");
        JMenuItem miPlaceholder1 = new JMenuItem("Export Magazzino (TODO)");
        miPlaceholder1.addActionListener(e -> log("TODO: integrare modulo export magazzino"));
        mFeatures.add(miPlaceholder1);

        JMenuItem miPlaceholder2 = new JMenuItem("Editor Articoli Easyfatt (TODO)");
        miPlaceholder2.addActionListener(e -> log("TODO: integrare EF Importer / editor articoli"));
        mFeatures.add(miPlaceholder2);

        mb.add(mFile);
        mb.add(mFeatures);
        mb.add(mView);
        return mb;
    }

    public void log(String s) {
        System.out.println(s);
        logWindow.append(s);
    }
}
