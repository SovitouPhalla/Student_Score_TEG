# 🎓 English Department: Student Report Portal

A streamlined web application designed to automate student term reports. This platform allows teachers to securely input scores and provides parents with a professional, real-time view of their child's academic progress.

## ✨ Key Features
* **Teacher Portal:** Secure login to input scores for specific classes (e.g., L10T4, L6T2(2)).
* **Weighted Grading:** Automatically calculates final grades based on Department standards:
    * **Conduct/CP:** 10% | **HW & Quizzes:** 30% | **Mid-Term:** 25% | **Final:** 35%
* **4-Term Tracking:** Supports a full academic year without overwriting previous data.
* **Parent Portal:** Simple Student ID & Password login to view formatted report cards.
* **Excel Database:** Uses `grades.xlsx` for easy manual backup and data portability.

## 🛠️ Tech Stack
* **Backend:** Python (Flask)
* **Database:** MS Excel (via Pandas & Openpyxl)
* **Frontend:** HTML5, CSS3, Jinja2 Templates

## 🚀 Quick Setup
1. **Clone the repo:**
   ```bash
   git clone [https://github.com/your-username/english-report-portal.git](https://github.com/your-username/english-report-portal.git)
