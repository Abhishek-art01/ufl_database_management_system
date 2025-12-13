# 🚖 Taxi Travel Management System

_A streamlined web application to automate corporate taxi bookings, employee travel tracking, and voucher generation using Streamlit and PostgreSQL._

---

## 📌 Table of Contents
- [Overview](#overview)
- [Business Problem](#business-problem)
- [Database & Schema](#database--schema)
- [Tools & Technologies](#tools--technologies)
- [Project Structure](#project-structure)
- [Key Features & Logic](#key-features--logic)
- [User Interface](#user-interface)
- [How to Run This Project](#how-to-run-this-project)
- [Future Enhancements](#future-enhancements)
- [Author & Contact](#author--contact)

---

## <a name="overview"></a>Overview
This project facilitates the digital management of corporate taxi travel. It replaces manual logbooks with a centralized database application that allows for instant trip identification, employee verification, bulk voucher generation, and historical record viewing. The system ensures data integrity through strict validation rules and duplicate checks.

---

## <a name="business-problem"></a>Business Problem
Managing employee transport manually leads to several operational inefficiencies:
- **Data Redundancy:** Duplicate voucher numbers and trip records.
- **Manual Errors:** Incorrect employee IDs or names entered during rush hours.
- **Lack of Visibility:** Difficulty in tracking past trips or auditing travel costs.
- **Time Consumption:** Manually filling details for multiple employees on the same trip is slow.

**Solution:** A unified interface to fetch trip details automatically, assign employees in bulk, and generate secure records in a PostgreSQL database.

---

## <a name="database--schema"></a>Database & Schema
The system connects to a **PostgreSQL** database containing the following key tables:

* **`application_data_dump`**: Raw data source for application-based trip requests.
* **`manual_data_dump`**: Raw data source for manual/ad-hoc trip requests.
* **`taxi_travels`**: The master transactional table where confirmed bookings are stored.
    * *Columns:* `trip_id`, `travel_date`, `employee_id`, `voucher_no`, `amount`, `direction`, etc.

---

## <a name="tools--technologies"></a>Tools & Technologies
- **Frontend:** Streamlit (Python)
- **Database:** PostgreSQL
- **Backend Logic:** Python (Pandas, Psycopg2)
- **Configuration:** TOML (Secrets Management)
- **Deployment:** Streamlit Community Cloud / Local Host

---

## <a name="project-structure"></a>Project Structure
The project follows a modular structure to separate data, logic, and configuration:

```text
project_root/
│
├── .streamlit/                    # Streamlit configuration
│   └── secrets.toml               # Database credentials (DO NOT UPLOAD TO GITHUB)
│
├── data/                          # Raw data directories
│   ├── application_files/         # Source for App-based trip data
│   └── manual_files/              # Source for Manual/Ad-hoc trip data
│
├── images/ 
│       ├──dashboard_preview.png   # UI screenshots and assets
│       ├──login_preview.png       # UI screenshots and assets
│       ├──trip_preview.png        # UI screenshots and assets
│
├── scripts/                       # Application source code
│   ├── data_loader_src.py         # ETL scripts to load dump data into DB
│   └── taxi_data_entry_web.py     # Main Streamlit application file
│
├── .gitignore                     # Git ignore rules (secrets, venv, etc.)
├── readme.md                      # Project documentation
└── requirements.txt               # Python dependencies
```

## <a name="key-features--logic"></a>Key Features & Logic

### 1. Hybrid Search Logic
* **Smart Lookup:** Automatically detects if a search is for an "Application" or "Manual" trip.
* **Strict Validation:** Enforces 7-digit Trip IDs for application searches to prevent bad data entry.
* **Auto-Fill:** Fetches Employee Name, Gender, and Address instantly, eliminating manual typing.

### 2. Intelligent Voucher Management
* **Uniqueness Check:** It automatically generates unique voucher codes with an incremental serial number, enabling efficient sorting and structured data management.

* **Smart Suffixing:** If multiple employees are selected for one trip (e.g., Carpooling), the system automatically splits the voucher (e.g., `9001A`, `9001B`) while keeping the base number linked.

### 3. Bulk Operations
* **Multi-Select Interface:** A checkbox-enabled table allows the user to select specific employees from a group for a single trip.
* **Batch Save:** Writes multiple records (one per employee) to the database in a single transaction.

---

## <a name="user-interface"></a>User Interface

The application utilizes a compact "Single Page" design to maximize efficiency for power users:

* **Tabbed Navigation:**
    * **📝 Entry Tab:** The main workspace for processing new requests. It splits the screen between "Search Results" and "Entry Form" to avoid scrolling.
    * **📊 Records Tab:** A read-only view of historical data with filtering capabilities.
* **Dynamic Forms:** The search bar changes automatically based on the selected "Travel Type" (Manual vs Application), showing or hiding date fields as needed.

---

## <a name="how-to-run-this-project"></a>How to Run This Project

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Abhishek-art01/Taxi_management_db.git
   cd Taxi_management_db

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt


3. **Setup Secrets:**
   * Create a folder named .streamlit in the root directory.
   * Create a file named secrets.toml inside it.
   * Add your PostgreSQL credentials:

```toml
[postgres]
host = "your-db-host"
dbname = "your-db-name"
user = "your-user"
password = "your-password"
port = 5432
```

4. **Run the App:**
   ```bash
   streamlit run scripts/taxi_data_entry_webapp.py


## <a name="future-roadmap"></a>  Future Roadmap
We plan to scale this system with the following enhancements:

📊 Analytics Dashboard: Add a visual layer to track Total Spend per Vendor, Peak Shift Times, and Cost per Employee.

📱 Mobile Employee View: A mobile-friendly page where employees can upload their trip details and Taxi Bill to claim their trip cost directly.


📧 Automated Notifications: Trigger email or SMS alerts to employees containing their Voucher details immediately after payment intiated.

🔐 Role-Based Access Control (RBAC): Implement login levels (Admin vs. Viewer) to secure sensitive data.

---

## <a name="author--contact"></a>Author & Contact
Abhishek Pandey (Data Analyst)

📧 Email: abhishekpandey4577@gmain.com

🔗 LinkedIn: www.linkedin.com/in/abhishek-art01 