package app;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;

public class Config {

    private String dbUrl;
    private String dbUser;
    private String dbPassword;

    public Config() {
        loadConfig();
    }

    private void loadConfig() {
        try (InputStream is = getClass().getClassLoader().getResourceAsStream("config.json")) {

            String content = new String(is.readAllBytes(), StandardCharsets.UTF_8);
            
            this.dbUrl = extractJsonValue(content, "db_url");
            this.dbUser = extractJsonValue(content, "db_user");
            this.dbPassword = extractJsonValue(content, "db_password");

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        }
    }

    private String extractJsonValue(String json, String key) {
        String searchKey = "\"" + key + "\"";
        int keyIndex = json.indexOf(searchKey);
        if (keyIndex == -1) return "";
        
        int startIndex = json.indexOf("\"", keyIndex + searchKey.length() + 1) + 1;
        int endIndex = json.indexOf("\"", startIndex);
        return json.substring(startIndex, endIndex);
    }

    public String getDbUrl() { return dbUrl; }
    public String getDbUser() { return dbUser; }
    public String getDbPassword() { return dbPassword; }
}