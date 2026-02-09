package core.config;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

/**
 * Minimal config helper for config.properties placed in project root.
 */
public class ConfigUtils {

    public static Properties load(File file) throws IOException {
        Properties p = new Properties();
        if (file.exists()) {
            try (FileInputStream in = new FileInputStream(file)) {
                p.load(in);
            }
        }
        return p;
    }

    public static void save(File file, Properties p, String comments) throws IOException {
        try (FileOutputStream out = new FileOutputStream(file)) {
            p.store(out, comments);
        }
    }

    public static String get(Properties p, String key, String def) {
        String v = p.getProperty(key);
        return (v == null) ? def : v;
    }

    public static void set(Properties p, String key, String value) {
        if (value == null) p.remove(key);
        else p.setProperty(key, value);
    }
}
