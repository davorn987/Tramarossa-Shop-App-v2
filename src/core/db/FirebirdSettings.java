package core.db;

/**
 * Firebird connection settings.
 *
 * Supports:
 * - embedded by dbPath (ex: C:\db\file.eft)
 * - direct JDBC URL override (ex: jdbc:firebirdsql:embedded:C:\db\file.eft)
 *
 * If jdbcUrlOverride is not empty, it is used as-is.
 */
public class FirebirdSettings {
    private String dbPath;            // for embedded (path to .eft/.fdb)
    private String jdbcUrlOverride;   // if set, overrides dbPath mode
    private String user;
    private String password;
    private String jnaLibraryPath;    // optional (if you need to point to fbclient)

    public String getDbPath() { return dbPath; }
    public void setDbPath(String dbPath) { this.dbPath = dbPath; }

    public String getJdbcUrlOverride() { return jdbcUrlOverride; }
    public void setJdbcUrlOverride(String jdbcUrlOverride) { this.jdbcUrlOverride = jdbcUrlOverride; }

    public String getUser() { return user; }
    public void setUser(String user) { this.user = user; }

    public String getPassword() { return password; }
    public void setPassword(String password) { this.password = password; }

    public String getJnaLibraryPath() { return jnaLibraryPath; }
    public void setJnaLibraryPath(String jnaLibraryPath) { this.jnaLibraryPath = jnaLibraryPath; }
}
