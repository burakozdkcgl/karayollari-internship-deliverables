package app;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class Database {

    // Singleton Instance
    private static Database instance;
    private Connection connection;
    private Config config;

    // Private constructor
    private Database() {
        this.config = new Config();
    }

    // Singleton Instance
    public static synchronized Database getInstance() {
        if (instance == null) {
            instance = new Database();
        }
        return instance;
    }

    public Connection connect() {
        try {
            if (connection == null || connection.isClosed()) {
                connection = DriverManager.getConnection(
                    config.getDbUrl(),
                    config.getDbUser(),
                    config.getDbPassword()
                );
                System.out.println("[Database] Successfully connected to the database.");
            }
        } catch (SQLException e) {
            System.err.println(e.getMessage());
        }
        return connection;
    }

    public Connection getConnection() {
        try {
            if (connection == null || connection.isClosed()) {
                return connect();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return connection;
    }

    public void disconnect() {
        try {
            if (connection != null && !connection.isClosed()) {
                connection.close();
            }
        } catch (SQLException e) {}
    }
}
