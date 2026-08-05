package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import org.springframework.stereotype.Component;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

@Component
public class MockData {

    private final Auth auth;

    @PersistenceContext
    private EntityManager entityManager;

    public MockData(Auth auth) {
        this.auth = auth;
    }

    @Transactional
    public void insertMockData() {
        // Users
        auth.registerUser("admin", "System Administrator", "admin", "ADMIN");
        auth.registerUser("ahmet", "Ahmet Yilmaz", "123", "PERSONNEL");
        auth.registerUser("mehmet", "Mehmet Demir", "123", "PERSONNEL");

        // Fetch initialized master data from database
        List<FleetStatus> statuses = entityManager.createQuery("SELECT s FROM FleetStatus s", FleetStatus.class).getResultList();
        
        // Regions list to populate
        List<Region> regions = new ArrayList<>();
        regions.add(new Region("Section 1 - Bolu Pass"));
        regions.add(new Region("Section 2 - Izmir Highway"));
        regions.add(new Region("Section 4 - Ankara Ring Road"));
        regions.add(new Region("Section 10 - Trabzon Zigana"));
        regions.add(new Region("Section 14 - Bursa Site"));
        regions.add(new Region("Section 6 - Antalya Coastal Road"));

        for (Region region : regions) {
            entityManager.persist(region);
        }

        // Vehicles Data Generator (12 vehicles with random Region and Status)
        Random random = new Random();
        String[][] rawVehicles = {
            {"34 KGM 001", "Snowplow"},
            {"06 KGM 102", "Asphalt Paver"},
            {"35 KGM 405", "Patrol Vehicle"},
            {"16 KGM 708", "Crawler Dozer"},
            {"61 KGM 901", "Rig & Crane"},
            {"07 KGM 332", "Excavator"},
            {"41 KGM 554", "Dump Truck"},
            {"01 KGM 882", "Road Roller"},
            {"55 KGM 219", "Patrol Vehicle"},
            {"26 KGM 640", "Snowplow"},
            {"10 KGM 118", "Tractor Loader"},
            {"38 KGM 975", "Asphalt Paver"}
        };

        for (String[] v : rawVehicles) {
            FleetStatus randomStatus = statuses.get(random.nextInt(statuses.size()));
            Region randomRegion = regions.get(random.nextInt(regions.size()));
            
            entityManager.persist(new Vehicle(v[0], v[1], randomStatus, randomRegion));
        }

        System.out.println("[Database] Mock data inserted successfully with random assignments.");
    }
}