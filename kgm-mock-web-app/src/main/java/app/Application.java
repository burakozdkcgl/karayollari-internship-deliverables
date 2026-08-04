package app;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

import java.util.Properties;

@SpringBootApplication
public class Application {

    public static void main(String[] args) {
        Config config = new Config();

        Properties props = System.getProperties();
        props.put("spring.datasource.url", config.getDbUrl());
        props.put("spring.datasource.username", config.getDbUser());
        props.put("spring.datasource.password", config.getDbPassword());

        if (config.isResetSchemas()) {
            props.put("spring.jpa.hibernate.ddl-auto", "create");
        } else {
            props.put("spring.jpa.hibernate.ddl-auto", "update");
        }

        SpringApplication.run(Application.class, args);
    }

    @Bean
    public CommandLineRunner initDatabase(MockData mockData) {
        return args -> {
            Config config = new Config();
            if (config.isResetSchemas() && config.isIfResetFillMockData()) {
                mockData.insertMockData();
            }
        };
    }
}