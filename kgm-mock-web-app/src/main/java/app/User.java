package app;

import jakarta.persistence.*;

@Entity
@Table(name = "users")
public class User {
    
    public enum Role {
        ADMIN,
        PERSONNEL
    }

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "user_id")
    private Long userId;

    @Column(name = "user_username", nullable = false, unique = true, length = 50)
    private String userUsername;

    @Column(name = "user_fullname", nullable = false, length = 100)
    private String userFullname;

    @Column(name = "user_password", nullable = false)
    private String userPassword;

    @Enumerated(EnumType.STRING)
    @Column(name = "user_role", nullable = false)
    private Role userRole;

    public User() {}

    public User(String userUsername, String userFullname, String userPassword, Role userRole) {
        this.userUsername = userUsername;
        this.userFullname = userFullname;
        this.userPassword = userPassword;
        this.userRole = userRole;
    }

    public Long getUserId() { return userId; }
    public void setUserId(Long userId) { this.userId = userId; }

    public String getUserUsername() { return userUsername; }
    public void setUserUsername(String userUsername) { this.userUsername = userUsername; }

    public String getUserFullname() { return userFullname; }
    public void setUserFullname(String userFullname) { this.userFullname = userFullname; }

    public String getUserPassword() { return userPassword; }
    public void setUserPassword(String userPassword) { this.userPassword = userPassword; }

    public Role getUserRole() { return userRole; }
    public void setUserRole(Role userRole) { this.userRole = userRole; }
}