package core.ui;

import features.shop.ui.ShopAppFrame;

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
        center.add(new JLabel("Tramarossa ToolSuite - scegli una funzione dal menu", SwingConstants.CENTER), BorderLayout.CENTER);
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
        JMenuItem miShop = new JMenuItem("Shop: Export/Statistiche/Venduto...");
        miShop.addActionListener(e -> new ShopAppFrame().showUI());
        mFeatures.add(miShop);

        JMenuItem miEf = new JMenuItem("Easyfatt: Editor Articoli... (TODO)");
        miEf.addActionListener(e -> log("TODO: integrare modulo EF-Importer"));
        mFeatures.add(miEf);

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
