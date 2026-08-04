package app;

import org.springframework.stereotype.Component;
import org.springframework.transaction.annotation.Transactional;

@Component
public class MockData {

    private final Auth auth;

    public MockData(Auth auth) {
        this.auth = auth;
    }

    @Transactional
    public void insertMockData() {
        auth.registerUser("admin", "Sistem Yöneticisi", "admin", User.Role.ADMIN);
        auth.registerUser("ahmet", "Ahmet Yılmaz", "123", User.Role.PERSONNEL);
        auth.registerUser("mehmet", "Mehmet Demir", "123", User.Role.PERSONNEL);
        auth.registerUser("ayse", "Ayşe Kaya", "123", User.Role.PERSONNEL);

        System.out.println("[Database] Mock data inserted successfully.");
    }
}