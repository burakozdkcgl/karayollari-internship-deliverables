package app;

import jakarta.persistence.*;
import jakarta.servlet.http.HttpSession;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

// ==========================================
// 1. ENTITIES (Database Tables)
// ==========================================

@Entity
@Table(name = "regions")
class Region {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "region_name", nullable = false, unique = true, length = 100)
    private String name;

    public Region() {}
    public Region(String name) { this.name = name; }

    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
}

@Entity
@Table(name = "fleet_statuses")
class FleetStatus {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "status_code", nullable = false, unique = true, length = 30)
    private String statusCode; // ACTIVE, PASSIVE, MAINTENANCE

    public FleetStatus() {}
    public FleetStatus(String statusCode) { this.statusCode = statusCode; }

    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    public String getStatusCode() { return statusCode; }
    public void setStatusCode(String statusCode) { this.statusCode = statusCode; }
}

@Entity
@Table(name = "fleet_vehicles")
class Vehicle {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "vehicle_id")
    private Long id;

    @Column(name = "plate_number", nullable = false, unique = true, length = 20)
    private String plateNumber;

    @Column(name = "vehicle_type", nullable = false, length = 50)
    private String type;

    @ManyToOne
    @JoinColumn(name = "status_id", nullable = false)
    private FleetStatus status;

    @ManyToOne
    @JoinColumn(name = "region_id", nullable = false)
    private Region region;

    public Vehicle() {}

    public Vehicle(String plateNumber, String type, FleetStatus status, Region region) {
        this.plateNumber = plateNumber;
        this.type = type;
        this.status = status;
        this.region = region;
    }

    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    public String getPlateNumber() { return plateNumber; }
    public void setPlateNumber(String plateNumber) { this.plateNumber = plateNumber; }
    public String getType() { return type; }
    public void setType(String type) { this.type = type; }
    public FleetStatus getStatus() { return status; }
    public void setStatus(FleetStatus status) { this.status = status; }
    public Region getRegion() { return region; }
    public void setRegion(Region region) { this.region = region; }
}

// ==========================================
// 2. CONTROLLER & API ENDPOINTS
// ==========================================

@Controller
@RequestMapping("/api/fleet")
public class Fleet {

    @PersistenceContext
    private EntityManager entityManager;

    private boolean isAuthenticated(HttpSession session) {
        return session.getAttribute("user") != null;
    }

    @GetMapping
    @ResponseBody
    public ResponseEntity<List<Vehicle>> getAllVehicles(HttpSession session) {
        if (!isAuthenticated(session)) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        List<Vehicle> vehicles = entityManager.createQuery(
                "SELECT v FROM Vehicle v JOIN FETCH v.status JOIN FETCH v.region ORDER BY v.id DESC", Vehicle.class)
                .getResultList();
        return ResponseEntity.ok(vehicles);
    }

    @GetMapping("/regions")
    @ResponseBody
    public ResponseEntity<List<Region>> getRegions(HttpSession session) {
        if (!isAuthenticated(session)) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        List<Region> regions = entityManager.createQuery("SELECT r FROM Region r", Region.class).getResultList();
        return ResponseEntity.ok(regions);
    }

    @GetMapping("/statuses")
    @ResponseBody
    public ResponseEntity<List<FleetStatus>> getStatuses(HttpSession session) {
        if (!isAuthenticated(session)) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        List<FleetStatus> statuses = entityManager.createQuery("SELECT s FROM FleetStatus s", FleetStatus.class).getResultList();
        return ResponseEntity.ok(statuses);
    }

    @PutMapping("/{id}")
    @Transactional
    @ResponseBody
    public ResponseEntity<Vehicle> updateVehicleStatusAndRegion(
            @PathVariable Long id,
            @RequestParam Long statusId,
            @RequestParam Long regionId,
            HttpSession session) {

        if (!isAuthenticated(session)) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();

        Vehicle vehicle = entityManager.find(Vehicle.class, id);
        FleetStatus status = entityManager.find(FleetStatus.class, statusId);
        Region region = entityManager.find(Region.class, regionId);

        if (vehicle == null || status == null || region == null) {
            return ResponseEntity.badRequest().build();
        }

        vehicle.setStatus(status);
        vehicle.setRegion(region);
        entityManager.merge(vehicle);

        return ResponseEntity.ok(vehicle);
    }

    @GetMapping("/stats")
    @ResponseBody
    public ResponseEntity<Map<String, Long>> getFleetStatusStats(HttpSession session) {
        if (!isAuthenticated(session)) return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();

        List<Object[]> results = entityManager.createQuery(
                "SELECT v.status.statusCode, COUNT(v) FROM Vehicle v GROUP BY v.status.statusCode", Object[].class)
                .getResultList();

        Map<String, Long> stats = new HashMap<>();
        for (Object[] result : results) {
            stats.put((String) result[0], (Long) result[1]);
        }
        return ResponseEntity.ok(stats);
    }
}