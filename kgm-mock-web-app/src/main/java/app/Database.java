package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import org.springframework.stereotype.Component;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

@Component
public class Database {

    @PersistenceContext
    private EntityManager entityManager;

    @Transactional
    public void initMasterData() {
        initUserRoles();
        initFleetStatuses();
    }

    private void initUserRoles() {
        List<UserRole> roles = entityManager.createQuery("SELECT r FROM UserRole r", UserRole.class).getResultList();
        if (roles.isEmpty()) {
            entityManager.persist(new UserRole("ADMIN"));
            entityManager.persist(new UserRole("PERSONNEL"));
            System.out.println("[Database] Default user roles initialized.");
        }
    }

    private void initFleetStatuses() {
        List<FleetStatus> statuses = entityManager.createQuery("SELECT s FROM FleetStatus s", FleetStatus.class).getResultList();
        if (statuses.isEmpty()) {
            entityManager.persist(new FleetStatus("ACTIVE"));
            entityManager.persist(new FleetStatus("PASSIVE"));
            entityManager.persist(new FleetStatus("MAINTENANCE"));
            System.out.println("[Database] Default fleet statuses initialized.");
        }
    }
}