package core.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * Single DB manager for Firebird.
 *
 * - If settings.jdbcUrlOverride is set => uses it as-is
 * - Else uses embedded URL from settings.dbPath
 *
 * Note: Jaybird driver must be on classpath.
 */
public class DbManager {
    private final FirebirdSettings settings;

    public DbManager(FirebirdSettings settings) {
        this.settings = settings;
    }

    public FirebirdSettings getSettings() {
        return settings;
    }

    public Connection getConnection() throws SQLException {
        // Optional: set jna.library.path before first connect (if needed)
        if (settings.getJnaLibraryPath() != null && !settings.getJnaLibraryPath().trim().isEmpty()) {
            // This must be set before JNA loads native libs; keeping it here for simplicity
            System.setProperty("jna.library.path", settings.getJnaLibraryPath().trim());
        }

        // Load Jaybird driver (safe to call multiple times)
        try {
            Class.forName("org.firebirdsql.jdbc.FBDriver");
        } catch (ClassNotFoundException e) {
            throw new SQLException("Jaybird driver not found: org.firebirdsql.jdbc.FBDriver", e);
        }

        String url = buildJdbcUrl(settings);
        return DriverManager.getConnection(url, settings.getUser(), settings.getPassword());
    }

    private static String buildJdbcUrl(FirebirdSettings s) throws SQLException {
        String override = s.getJdbcUrlOverride();
        if (override != null && !override.trim().isEmpty()) {
            return override.trim();
        }

        String path = s.getDbPath();
        if (path == null || path.trim().isEmpty()) {
            throw new SQLException("Missing dbPath (and jdbcUrlOverride not set).");
        }

        // normalize backslashes to forward slashes for JDBC url
        String norm = path.trim().replace("\\", "/");
        return "jdbc:firebirdsql:embedded:" + norm;
    }
}
