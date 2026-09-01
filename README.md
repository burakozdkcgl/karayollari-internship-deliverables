# karayollari internship deliverables
tracking my summer internship deliverables at turkiye general directorate of highways

## Repository Overview

* `kgm-mock-web-app/` - a java spring boot mock web application simulated for a government institution, built as an educational case study.
* `excel-vba-case-tracker-app/` - automated case tracking system and a detailed LaTeX report of it
* `excel-vba-employee-management/` - interactive dashboard to query employee seniority and leave entitlement schedules
* `internship-report.pdf` - internship report document

### KGM Mock Web Application

#### Preview

| Landing & Authentication |
| :---: |
| ![Landing Page](./kgm-mock-web-app/screenshots/screenshot1.jpg) |
| *Portal entry and role-based authentication* |

| Operational Control Center | Road & Distance Routing |
| :---: | :---: |
| ![Overview](./kgm-mock-web-app/screenshots/screenshot2.jpg) | ![Distance](./kgm-mock-web-app/screenshots/screenshot3.jpg) |
| *Real-time 3-shift clock & telemetry* | *Leaflet.js & OSRM routing engine* |

| Fleet Operations Management | Personnel Directory |
| :---: | :---: |
| ![Fleet](./kgm-mock-web-app/screenshots/screenshot4.jpg) | ![Users](./kgm-mock-web-app/screenshots/screenshot5.jpg) |
| *Regional tracking & Chart.js statistics* | *BLOB image directory & role management* |

#### Quick Start

1. Set your database credentials in `kgm-mock-web-app/src/main/resources/config.json`:

```json
{
  "db_url": "jdbc:mysql://localhost:3306/kgm",
  "db_user": "root",
  "db_password": "password",
  "RESET_SCHEMAS_ON_EACH_LAUNCH": true,
  "IF_RESET_FILL_MOCK_DATA": true
}
```

2. Build and run:

```bash
cd kgm-mock-web-app
mvn spring-boot:run
```

3. Open `http://localhost:8080` in your browser.
   * Default credentials: `admin` for both username and password.
