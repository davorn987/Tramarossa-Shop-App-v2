package core.ui;

import javax.swing.*;
import java.awt.*;

public class LogWindow extends JFrame {
    private final JTextArea textArea = new JTextArea(25, 100);

    public LogWindow() {
        super("Log");
        setDefaultCloseOperation(WindowConstants.HIDE_ON_CLOSE);
        textArea.setEditable(false);
        setLayout(new BorderLayout());
        add(new JScrollPane(textArea), BorderLayout.CENTER);
        pack();
        setLocationRelativeTo(null);
    }

    public void append(String s) {
        SwingUtilities.invokeLater(() -> {
            textArea.append(s + "\n");
            textArea.setCaretPosition(textArea.getDocument().getLength());
        });
    }

    public void clear() {
        SwingUtilities.invokeLater(() -> textArea.setText(""));
    }
}
