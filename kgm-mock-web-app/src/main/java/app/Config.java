package app;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.InputStream;

public class Config {

    private String dbUrl = "";
    private String dbUser = "";
    private String dbPassword = "";
    private boolean resetSchemas;
    private boolean ifResetFillMockData;

    public Config() {
        loadConfig();
    }

    private void loadConfig() {
        ObjectMapper mapper = new ObjectMapper();
        try (InputStream is = getClass().getClassLoader().getResourceAsStream("config.json")) {
            if (is == null) return;
            
            JsonNode root = mapper.readTree(is);
            
            if (root.has("db_url")) this.dbUrl = root.get("db_url").asText("");
            if (root.has("db_user")) this.dbUser = root.get("db_user").asText("");
            if (root.has("db_password")) this.dbPassword = root.get("db_password").asText("");
            if (root.has("RESET_SCHEMAS_ON_EACH_LAUNCH")) {
                this.resetSchemas = root.get("RESET_SCHEMAS_ON_EACH_LAUNCH").asBoolean(false);
            }
            if (root.has("IF_RESET_FILL_MOCK_DATA")) {
                this.ifResetFillMockData = root.get("IF_RESET_FILL_MOCK_DATA").asBoolean(false);
            }

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    public String getDbUrl() { return dbUrl; }
    public String getDbUser() { return dbUser; }
    public String getDbPassword() { return dbPassword; }
    public boolean isResetSchemas() { return resetSchemas; }
    public boolean isIfResetFillMockData() { return ifResetFillMockData; }
}