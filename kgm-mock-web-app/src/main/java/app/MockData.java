package app;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import org.springframework.stereotype.Component;
import org.springframework.transaction.annotation.Transactional;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import java.util.Base64;

@Component
public class MockData {

    private final Auth auth;

    @PersistenceContext
    private EntityManager entityManager;

    public MockData(Auth auth) {
        this.auth = auth;
    }

    private byte[] loadAvatar(String filename) {
        String path = "images/avatars/" + filename;
        try (InputStream is = getClass().getClassLoader().getResourceAsStream(path)) {
            if (is != null) {
                return is.readAllBytes();
            }
        } catch (Exception e) {
            System.err.println("Error on " + path);
        }
        return null;
    }

    @Transactional
    public void insertMockData() {
        auth.registerUser("admin", "Cazibe Abbasova", "cazibe.abbasova@kgm.gov.tr", loadAvatar("cazibe.jpg"), "admin", "ADMIN");
        auth.registerUser("simge", "Simge Kunter", "simge.paris@kgm.gov.tr", loadAvatar("simge.jpg"), "123", "PERSONNEL");
        auth.registerUser("kemal", "Kemal Kükreyen", "kemal.kukreyen@kgm.gov.tr", loadAvatar("kemal.jpg"), "123", "PERSONNEL");
        auth.registerUser("baris", "Barış Martı", "baris.marti@kgm.gov.tr", loadAvatar("baris.jpg"), "123", "PERSONNEL");
        auth.registerUser("peker", "Peker Pekmez", "peker.pekmez@kgm.gov.tr", loadAvatar("peker.jpg"), "123", "ADMIN");

        // Master Data
        List<FleetStatus> statuses = entityManager.createQuery("SELECT s FROM FleetStatus s", FleetStatus.class).getResultList();
        
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

        // Vehicles
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

        System.out.println("[Database] Mock data inserted successfully.");
    }
}