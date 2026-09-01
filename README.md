# karayollari internship deliverables
tracking my summer internship deliverables at turkiye general directorate of highways

## Repository Overview

* `kgm-mock-web-app/` - a java spring boot mock web application simulated for a government institution, built as an educational case study.
* `excel-vba-case-tracker-app/` - automated case tracking system and a detailed LaTeX report of it
* `excel-vba-employee-management/` - interactive dashboard to query employee seniority and leave entitlement schedules
* `internship-report.pdf` - internship report document

## KGM Mock Web Application

### Preview

| Landing Page |
| :---: |
| ![Landing Page](./kgm-mock-web-app/screenshots/screenshot1.jpg) |
| *login screen with role-based auth* |

| Main Dashboard | Distance Calculator |
| :---: | :---: |
| ![Overview](./kgm-mock-web-app/screenshots/screenshot2.jpg) | ![Distance](./kgm-mock-web-app/screenshots/screenshot3.jpg) |
| *live 3-shift clock and system status* | *route and distance calculation between cities* |

| Fleet Management | Users Directory |
| :---: | :---: |
| ![Fleet](./kgm-mock-web-app/screenshots/screenshot4.jpg) | ![Users](./kgm-mock-web-app/screenshots/screenshot5.jpg) |
| *vehicle status tracking and chart breakdown* | *registered personnel list* |

### How to run

1. update db credentials in `kgm-mock-web-app/src/main/resources/config.json`

```json
{
  "db_url": "jdbc:mysql://localhost:3306/kgm",
  "db_user": "root",
  "db_password": "password",
  "RESET_SCHEMAS_ON_EACH_LAUNCH": true,
  "IF_RESET_FILL_MOCK_DATA": true
}
```

2. build and run

```bash
cd kgm-mock-web-app
mvn spring-boot:run
```

3. visit `http://localhost:8080`
   * default credentials are `admin` for both username and password

## Excel VBA Apps

> note: the original `.xlsm` workbooks are not included in this repository due to confidentiality policies. you cannot directly run these scripts without the internal sheet structure; they are kept here purely for version control and backup purposes.

### 1. Case Tracker App (`excel-vba-case-tracker-app/`)

automated case tracking and management system. the folder also includes a comprehensive technical report written in LaTeX covering system architecture, user workflows, and maintenance guidelines for future developers.

| | | |
| :---: | :---: | :---: |
| <img src="./excel-vba-case-tracker-app/report-source-code/screenshot2.jpg" style="max-height: 180px; width: auto;" /> | <img src="./excel-vba-case-tracker-app/report-source-code/screenshot3.jpg" style="max-height: 180px; width: auto;" /> | <img src="./excel-vba-case-tracker-app/report-source-code/screenshot4.jpg" style="max-height: 180px; width: auto;" /> |

### 2. Employee Management Dashboard (`excel-vba-employee-management/`)

an interactive vba tool to query employee seniority, employment dates, and annual leave entitlement schedules by ID.

| | | |
| :---: | :---: | :---: |
| <img src="./excel-vba-employee-management/screenshot1.jpg" style="max-height: 180px; width: auto;" /> | <img src="./excel-vba-employee-management/screenshot2.jpg" style="max-height: 180px; width: auto;" /> | <img src="./excel-vba-employee-management/screenshot3.jpg" style="max-height: 180px; width: auto;" /> |