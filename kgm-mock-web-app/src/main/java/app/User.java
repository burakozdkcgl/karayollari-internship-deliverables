package app;

import jakarta.persistence.*;

@Entity
@Table(name = "user_roles")
class UserRole {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "role_id")
    private Long id;

    @Column(name = "role_name", nullable = false, unique = true, length = 30)
    private String name;

    public UserRole() {}

    public UserRole(String name) {
        this.name = name;
    }

    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
}

@Entity
@Table(name = "users")
public class User {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "user_id")
    private Long userId;

    @Column(name = "user_username", nullable = false, unique = true, length = 50)
    private String userUsername;

    @Column(name = "user_fullname", nullable = false, length = 100)
    private String userFullname;

    @Column(name = "user_email", length = 100)
    private String userEmail;

    @Lob
    @Column(name = "profile_image", columnDefinition = "LONGBLOB")
    private byte[] profileImage;

    @Column(name = "user_password", nullable = false)
    private String userPassword;

    @ManyToOne
    @JoinColumn(name = "role_id", nullable = false)
    private UserRole role;

    public User() {}

    public User(String userUsername, String userFullname, String userEmail, byte[] profileImage, String userPassword, UserRole role) {
        this.userUsername = userUsername;
        this.userFullname = userFullname;
        this.userEmail = userEmail;
        this.profileImage = profileImage;
        this.userPassword = userPassword;
        this.role = role;
    }

    public Long getUserId() { return userId; }
    public void setUserId(Long userId) { this.userId = userId; }

    public String getUserUsername() { return userUsername; }
    public void setUserUsername(String userUsername) { this.userUsername = userUsername; }

    public String getUserFullname() { return userFullname; }
    public void setUserFullname(String userFullname) { this.userFullname = userFullname; }

    public String getUserEmail() { return userEmail; }
    public void setUserEmail(String userEmail) { this.userEmail = userEmail; }

    public byte[] getProfileImage() { return profileImage; }
    public void setProfileImage(byte[] profileImage) { this.profileImage = profileImage; }

    public String getUserPassword() { return userPassword; }
    public void setUserPassword(String userPassword) { this.userPassword = userPassword; }

    public UserRole getRole() { return role; }
    public void setRole(UserRole role) { this.role = role; }
}