package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import jakarta.servlet.http.HttpSession;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
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
                if (hex.length() == 1) hexString.append('0');
                hexString.append(hex);
            }
            return hexString.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("ERROR_HASHING_PASSWORD", e);
        }
    }

    @Transactional
    public void registerUser(String username, String fullName, String email, byte[] profileImage, String rawPassword, String roleName) {
        String hashedPassword = hashPassword(rawPassword);
        UserRole userRole = entityManager.createQuery("SELECT r FROM UserRole r WHERE r.name = :roleName", UserRole.class)
                .setParameter("roleName", roleName)
                .getSingleResult();

        User user = new User(username, fullName, email, profileImage, hashedPassword, userRole);
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

    @GetMapping("/")
    public String indexPage(HttpSession session) {
        return session.getAttribute("user") != null ? "redirect:/verified" : "redirect:/index";
    }

    @GetMapping("/index")
    public String index(HttpSession session) {
        return session.getAttribute("user") != null ? "redirect:/verified" : "index";
    }

    @GetMapping("/verified")
    public String verified(HttpSession session) {
        return session.getAttribute("user") == null ? "redirect:/index" : "verified";
    }

    @PostMapping("/login")
    public String login(@RequestParam String username, @RequestParam String password, HttpSession session) {
        User user = authenticate(username, password);
        if (user != null) {
            session.setAttribute("user", user);
            return "redirect:/verified";
        }
        return "redirect:/index?error=INVALID_CREDENTIALS";
    }

    @GetMapping("/logout")
    public String logout(HttpSession session) {
        session.invalidate();
        return "redirect:/index";
    }

    @GetMapping("/api/user-info")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> getUserInfo(HttpSession session) {
        User user = (User) session.getAttribute("user");
        if (user == null) {
            return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        }
        Map<String, Object> data = new HashMap<>();
        data.put("id", user.getUserId());
        data.put("name", user.getUserFullname());
        data.put("role", user.getRole() != null ? user.getRole().getName().toLowerCase() : "personnel");
        return ResponseEntity.ok(data);
    }

    @GetMapping("/api/users")
    @ResponseBody
    public ResponseEntity<List<User>> getAllUsers(HttpSession session) {
        if (session.getAttribute("user") == null) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        List<User> users = entityManager.createQuery("SELECT u FROM User u JOIN FETCH u.role", User.class).getResultList();
        return ResponseEntity.ok(users);
    }

    @GetMapping("/api/users/{id}/avatar")
    @ResponseBody
    public ResponseEntity<byte[]> getUserAvatar(@PathVariable Long id) {
        User user = entityManager.find(User.class, id);
        if (user != null && user.getProfileImage() != null) {
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.IMAGE_JPEG_VALUE)
                    .body(user.getProfileImage());
        }
        return ResponseEntity.notFound().build();
    }

    @PostMapping("/api/users/{id}/avatar")
    @Transactional
    @ResponseBody
    public ResponseEntity<String> uploadAvatar(@PathVariable Long id, @RequestParam("file") MultipartFile file, HttpSession session) throws IOException {
        if (session.getAttribute("user") == null) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        User user = entityManager.find(User.class, id);
        if (user == null) return ResponseEntity.notFound().build();

        user.setProfileImage(file.getBytes());
        entityManager.merge(user);
        return ResponseEntity.ok("OK");
    }
}