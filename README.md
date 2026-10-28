# Intelligent Indian Road Safety Platform (SafeRoad AI)

![Platform](https://img.shields.io/badge/SafeRoad_AI-India-10b981?style=for-the-badge&logo=shield)
![FastAPI](https://img.shields.io/badge/Backend-FastAPI_0.115-009688?style=flat&logo=fastapi)
![React](https://img.shields.io/badge/Frontend-React_18_Vite_Tailwind-61dafb?style=flat&logo=react)
![Scikit-Learn](https://img.shields.io/badge/ML-Scikit--learn_RandomForest-f7931e?style=flat&logo=scikitlearn)
![OpenCV](https://img.shields.io/badge/Vision-OpenCV_YOLO_Interface-5c3ee8?style=flat&logo=opencv)
![License](https://img.shields.io/badge/License-MIT-blue.svg?style=flat)

**SafeRoad AI** is an intelligent, full-stack road safety management and accident-prevention platform engineered specifically for Indian national highways, state highways, and congested urban corridors. It unifies machine learning collision prediction, computer vision pothole detection, interactive geospatial hazard mapping, emergency trauma hospital dispatch, and civic agency resolution workflows into a cohesive, responsive dashboard.

---

## 1. System Architecture

```
                                +-----------------------------------+
                                |    React + Vite Single Page App   |
                                |  (Tailwind CSS, Leaflet, ChartJS) |
                                +-----------------+-----------------+
                                                  |
                                                  | HTTP / JSON / Multipart
                                                  v
                                +-----------------+-----------------+
                                |        FastAPI REST API Engine    |
                                +-----------------+-----------------+
                                                  |
         +--------------------+-------------------+-------------------+--------------------+
         |                    |                   |                   |                    |
         v                    v                   v                   v                    v
  +--------------+    +---------------+   +---------------+   +---------------+    +---------------+
  | /api/predict |    | /api/hazards  |   | /api/potholes |   | /api/hospitals|    |  /api/routes  |
  +--------------+    +---------------+   +---------------+   +---------------+    +---------------+
         |                    |                   |                   |                    |
         v                    v                   v                   v                    v
  [Scikit-learn       [SQLAlchemy ORM     [OpenCV Vision      [Haversine Geo      [Corridor Hazard
   Random Forest]      SQLite/Postgres]    YOLO Bboxes]        Distance Calc]      Density Index]
```

---

## 2. Platform Feature Modules

1. **Safety Dashboard (`/`)**: Real-time corridor safety indices, hazard counters, monthly incident trend charts, category distributions, and active alert feeds.
2. **Interactive Safety Map (`/map`)**: Geocoded road hazard markers, critical blackspot danger zones, emergency hospital overlays, and severity filters for major Indian corridors (Bengaluru, Delhi-NCR, Mumbai, Chennai, Hyderabad).
3. **Accident-Risk Predictor (`/prediction`)**: ML-powered assessment taking Indian road parameters (road classification, monsoon downpours, surface waterlogging, lighting, speed delta, and vehicle mix) to output collision probability, top risk drivers, and preventive advisories.
4. **Pothole AI Vision (`/pothole-detection`)**: Computer vision interface analyzing road imagery with OpenCV and YOLO-format bounding box coordinates, crater depth estimation, and 1-click conversion into municipal hazard tickets.
5. **Hazard Reporting Portal (`/report`)**: Citizen and road marshal reporting form with GPS geocoding, severity grading, and immediate database registration.
6. **Route Safety Comparator (`/route-safety`)**: Side-by-side evaluation of alternate highway routes comparing safety indices, hazard densities, pothole frequencies, and lighting coverage.
7. **Emergency Hospital Locator (`/hospitals`)**: Golden-hour trauma directory with distance calculation, Level-1 center filtering, 24x7 ICU bed availability, and 108/112 direct call shortcuts.
8. **Admin Analytics & Governance (`/admin`)**: Municipal repair ticket workflows (BBMP, NHAI, PWD, MSRDC), city compliance scorecards, and live ML model performance telemetry.

---

## 3. Technology Stack

- **Frontend**:
  - React 18 & Vite
  - Tailwind CSS with customized road-safety palette
  - React Router DOM v7
  - React-Leaflet & Leaflet (Custom SVG markers, popup cards, danger radius circles)
  - Chart.js & react-chartjs-2 (Line trends, Doughnut categories, Bar time-of-day charts)
  - Axios API client with centralized error interception
  - Lucide React icon set
- **Backend**:
  - Python 3.11+ with FastAPI & Uvicorn
  - Pydantic v2 validation models
  - SQLAlchemy 2.0 ORM (configured for SQLite with seamless PostgreSQL migration)
  - Python-Multipart for file uploads
- **Machine Learning & Computer Vision**:
  - Scikit-learn Random Forest Classifier pipeline (`app/ml/accident_risk_model.joblib`)
  - Pandas & NumPy for feature pre-processing and calibration
  - OpenCV image processing, contour extraction, depth calculation, and normalized YOLO bounding box generator
- **Testing & Verification**:
  - Pytest test suite covering all API endpoints and validation constraints

---

## 4. Project Structure

```
safe_road_ai/
├── backend/
│   ├── app/
│   │   ├── api/routes/          # FastAPI routers (predict, hazards, potholes, hospitals, routes, analytics)
│   │   ├── core/                # Configuration and database connection
│   │   ├── ml/                  # Trained model artifacts & training scripts
│   │   ├── models/              # SQLAlchemy database models
│   │   ├── schemas/             # Pydantic v2 validation schemas
│   │   └── services/            # ML & CV inference services
│   ├── tests/                   # Pytest automated test suite
│   ├── requirements.txt         # Backend Python dependencies
│   └── run_backend.py           # Backend startup script (port 8001)
├── frontend/
│   ├── src/
│   │   ├── components/          # Reusable UI components (Navbar, AlertCard, StatCard)
│   │   ├── pages/               # 8 Core application pages
│   │   ├── services/            # Axios API client
│   │   ├── App.jsx              # Application router
│   │   └── index.css            # Tailwind design system
│   ├── package.json             # Frontend Node.js dependencies
│   └── vite.config.js           # Vite configuration with API proxy
├── report_images/               # High-res screenshots and architectural block diagrams
├── SAFEROAD_AI_PROJECT_REPORT.docx # Comprehensive academic project report
├── start.bat                    # Windows 1-click launcher
├── start.sh                     # Linux/macOS 1-click launcher
├── .gitignore                   # Production git ignore configuration
├── .gitattributes               # Line endings normalization
├── .editorconfig                # Universal indentation and charset rules
└── README.md                    # Project documentation
```

---

## 5. Local Setup & Execution Guide

### Prerequisites
- Node.js (v18+) and npm
- Python (v3.10+)

### One-Click Startup

- **Windows**: Double-click `start.bat` or run:
  ```cmd
  start.bat
  ```
- **Linux / macOS**: Run:
  ```bash
  chmod +x start.sh
  ./start.sh
  ```

---

### Manual Step-by-Step Setup

#### 1. Backend Setup
```bash
cd backend
pip install -r requirements.txt
python run_backend.py
```
- Server URL: `http://127.0.0.1:8001`
- Interactive Swagger API Docs: `http://127.0.0.1:8001/docs`

#### 2. Run Backend Pytest Tests
```bash
cd backend
python -m pytest
```

#### 3. Frontend Setup
```bash
cd frontend
npm install
npm run dev
```
- Dashboard URL: `http://localhost:3000`

---

## 6. Project Report & Academic Deliverable

An Anna University / Sri Eshwar College of Engineering academic project report replica has been created:
- **Report Document**: [`SAFEROAD_AI_PROJECT_REPORT.docx`](./SAFEROAD_AI_PROJECT_REPORT.docx)
- **Author**: Gopika S (Reg No: 722824104067)
- **Guide**: Dr. M. Suriya, M.E., Ph.D., Associate Professor
- **Course**: U23AD491 - Data Science
- **Batch**: 2024 - 2028 (October 2026)

---

## 7. License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.
