package core.net;

import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class FtpUtils {

    public static void uploadFileToFtp(String localFilePath, String remoteFileName,
                                       String server, String user, String pass, String remoteDir) throws IOException {
        FTPClient ftpClient = new FTPClient();
        try {
            ftpClient.connect(server);
            boolean login = ftpClient.login(user, pass);
            if (!login) {
                throw new IOException("Login FTP fallito");
            }
            ftpClient.enterLocalPassiveMode();
            ftpClient.setFileType(FTP.BINARY_FILE_TYPE);

            if (remoteDir != null && !remoteDir.trim().isEmpty()) {
                boolean changed = ftpClient.changeWorkingDirectory(remoteDir);
                if (!changed) {
                    throw new IOException("Impossibile cambiare directory remota: " + remoteDir);
                }
            }

            File localFile = new File(localFilePath);
            try (InputStream input = new FileInputStream(localFile)) {
                boolean ok = ftpClient.storeFile(remoteFileName, input);
                if (!ok) {
                    throw new IOException("Upload FTP fallito: " + ftpClient.getReplyString());
                }
            }

            ftpClient.logout();
        } finally {
            try { ftpClient.disconnect(); } catch (Exception ignore) {}
        }
    }

    public static void showFtpError(Component parent, Exception ex) {
        JOptionPane.showMessageDialog(parent,
                "Errore durante l'invio del file via FTP:\n" + ex.getMessage(),
                "Errore FTP",
                JOptionPane.ERROR_MESSAGE
        );
    }

    public static void showFtpSuccess(Component parent) {
        JOptionPane.showMessageDialog(parent,
                "File inviato con successo al server FTP.",
                "FTP Upload",
                JOptionPane.INFORMATION_MESSAGE
        );
    }
}
