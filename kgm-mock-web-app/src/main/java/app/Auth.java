package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.List;

@Service
public class Auth {

    @PersistenceContext
    private EntityManager entityManager;

    public static String hashPassword(String password) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] encodedhash = digest.digest(password.getBytes(StandardCharsets.UTF_8));
            StringBuilder hexString = new StringBuilder(2 * encodedhash.length);
            for (byte b : encodedhash) {
                String hex = Integer.toHexString(0xff & b);
                if (hex.length() == 1) {
                    hexString.append('0');
                }
                hexString.append(hex);
            }
            return hexString.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("Şifreleme algoritması bulunamadı!", e);
        }
    }

    @Transactional
    public void registerUser(String username, String fullName, String rawPassword, User.Role role) {
        String hashedPassword = hashPassword(rawPassword);
        User user = new User(username, fullName, hashedPassword, role);
        entityManager.persist(user);
    }

    public User authenticate(String username, String rawPassword) {
        String hashedPassword = hashPassword(rawPassword);
        List<User> users = entityManager.createQuery(
                "SELECT u FROM User u WHERE u.userUsername = :username AND u.userPassword = :password", User.class)
                .setParameter("username", username)
                .setParameter("password", hashedPassword)
                .getResultList();

        return users.isEmpty() ? null : users.get(0);
    }
}