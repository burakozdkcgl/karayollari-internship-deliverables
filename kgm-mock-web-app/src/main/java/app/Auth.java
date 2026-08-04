package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import jakarta.servlet.http.HttpSession;
import org.springframework.stereotype.Controller;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.List;

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
                if (hex.length() == 1) {
                    hexString.append('0');
                }
                hexString.append(hex);
            }
            return hexString.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("ERROR_HASHING_PASSWORD", e);
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

    @GetMapping("/")
    public String indexPage(HttpSession session) {
        if (session.getAttribute("user") != null) {
            return "redirect:/verified.html";
        }
        return "redirect:/index.html"; 
    }

    @GetMapping("/login")
    public String loginPage(HttpSession session) {
        if (session.getAttribute("user") != null) {
            return "redirect:/verified.html";
        }
        return "redirect:/index.html";
    }

    @PostMapping("/login")
    public String login(@RequestParam String username, 
                        @RequestParam String password, 
                        HttpSession session) {
        User user = authenticate(username, password);
        if (user != null) {
            session.setAttribute("user", user);
            return "redirect:/verified.html";
        }
        return "redirect:/index.html?error=INVALID_CREDENTIALS"; 
    }

    @GetMapping("/logout")
    public String logout(HttpSession session) {
        session.invalidate();
        return "redirect:/index.html";
    }

    @GetMapping("/api/user-info")
    @ResponseBody
    public String getUserInfo(HttpSession session) {
        User user = (User) session.getAttribute("user");
        if (user == null) {
            return "";
        }
        return user.getUserFullname();
    }
}