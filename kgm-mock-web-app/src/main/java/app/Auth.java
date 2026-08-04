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

    // Hash raw password using SHA-256
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

    // Register new user to the database
    @Transactional
    public void registerUser(String username, String fullName, String rawPassword, User.Role role) {
        String hashedPassword = hashPassword(rawPassword);
        User user = new User(username, fullName, hashedPassword, role);
        entityManager.persist(user);
    }

    // Validate credentials against stored user data
    public User authenticate(String username, String rawPassword) {
        String hashedPassword = hashPassword(rawPassword);
        List<User> users = entityManager.createQuery(
                "SELECT u FROM User u WHERE u.userUsername = :username AND u.userPassword = :password", User.class)
                .setParameter("username", username)
                .setParameter("password", hashedPassword)
                .getResultList();

        return users.isEmpty() ? null : users.get(0);
    }

    // Root endpoint: Redirects based on session state
    @GetMapping("/")
    public String indexPage(HttpSession session) {
        if (session.getAttribute("user") != null) {
            return "redirect:/verified";
        }
        return "redirect:/index"; 
    }

    // Login/Index endpoint: Block authenticated users from reaching login page
    @GetMapping("/index")
    public String index(HttpSession session) {
        if (session.getAttribute("user") != null) {
            return "redirect:/verified"; // Redirect to protected page if already logged in
        }
        return "index"; // Renders templates/index.html
    }

    // Protected endpoint: Restrict unauthenticated access
    @GetMapping("/verified")
    public String verified(HttpSession session) {
        if (session.getAttribute("user") == null) {
            return "redirect:/index"; // Redirect to login page if unauthenticated
        }
        return "verified"; // Renders templates/verified.html
    }

    // Process authentication request
    @PostMapping("/login")
    public String login(@RequestParam String username, 
                        @RequestParam String password, 
                        HttpSession session) {
        User user = authenticate(username, password);
        if (user != null) {
            session.setAttribute("user", user);
            return "redirect:/verified";
        }
        return "redirect:/index?error=INVALID_CREDENTIALS"; 
    }

    // Invalidate session and log out
    @GetMapping("/logout")
    public String logout(HttpSession session) {
        session.invalidate();
        return "redirect:/index";
    }

    // REST API endpoint to return active user details
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