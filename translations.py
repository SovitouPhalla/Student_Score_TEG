"""
translations.py
---------------
Hard-coded translation dictionary replacing Flask-Babel.

Usage (injected automatically via app.context_processor):

    {{ _('Student Name') }}                        → translated string
    {{ _('Passing threshold: %(pct)s%%', pct=50) }} → 'Passing threshold: 50%'

Add new keys to BOTH 'en' and 'km' blocks simultaneously.
English values are the canonical key text — they always match the key itself so
the fallback (returning the key unchanged) is invisible in the English UI.
"""

TRANSLATIONS: dict[str, dict[str, str]] = {
    # ──────────────────────────────────────────────────────────────────────────
    # ENGLISH
    # ──────────────────────────────────────────────────────────────────────────
    "en": {
        # ── Navigation ────────────────────────────────────────────────────────
        "Departments":                   "Departments",
        "Teacher Portal":                "Teacher Portal",
        "Admin":                         "Admin",
        "HOD Approvals":                 "HOD Approvals",
        "Teacher Logout":                "Teacher Logout",
        "Admin Logout":                  "Admin Logout",
        "Logout":                        "Logout",

        # ── Site-wide ─────────────────────────────────────────────────────────
        "Student Term Report Portal":    "Student Term Report Portal",
        "All rights reserved.":          "All rights reserved.",

        # ── Landing page ──────────────────────────────────────────────────────
        "Welcome to the Student Report Portal":
            "Welcome to the Student Report Portal",
        "Select your department below to view a term report card.":
            "Select your department below to view a term report card.",
        "English Department":            "English Department",
        "Chinese Department":            "Chinese Department",

        # ── Login (shared between EN and CN login pages) ──────────────────────
        "Parent / Guardian Login":       "Parent / Guardian Login",
        "Chinese Department Login":      "Chinese Department Login",
        "Enter your child's full name and the password provided by the school.":
            "Enter your child's full name and the password provided by the school.",
        "Student Full Name":             "Student Full Name",
        "Child's Class (to clarify)":    "Child's Class (to clarify)",
        "— Select Class —":              "— Select Class —",
        "Multiple students share this name. Please select your child's class to continue.":
            "Multiple students share this name. Please select your child's class to continue.",
        "Multiple students share this name. Please select your child's class.":
            "Multiple students share this name. Please select your child's class.",
        "Parent Password":               "Parent Password",
        "Enter password":                "Enter password",
        "View Report Card":              "View Report Card",
        "Wrong department?":             "Wrong department?",

        # ── Report card — school header ───────────────────────────────────────
        "Official Term Report Card":     "Official Term Report Card",
        "Academic Year 2024 \u2013 2025":
            "Academic Year 2024 \u2013 2025",
        "Chinese Department \u2013 Official Term Report Card":
            "Chinese Department \u2013 Official Term Report Card",
        "CN Program \u2022 Academic Year 2024 \u2013 2025":
            "CN Program \u2022 Academic Year 2024 \u2013 2025",

        # ── Report card — student meta block ─────────────────────────────────
        "Student Name":                  "Student Name",
        "Student ID":                    "Student ID",
        "No. in List":                   "No. in List",
        "Class":                         "Class",
        "Department":                    "Department",
        "Chinese (CN Program)":          "Chinese (CN Program)",
        "Report Generated":              "Report Generated",
        "Terms Released":                "Terms Released",

        # ── Report card — term selector ───────────────────────────────────────
        "View Term":                     "View Term",
        "Master Overview (All Terms)":   "Master Overview (All Terms)",
        "Term %(n)s":                    "Term %(n)s",
        "Term %(n)s Report":             "Term %(n)s Report",
        "Term %(n)s Report \u2013 Chinese Department":
            "Term %(n)s Report \u2013 Chinese Department",

        # ── Report card — score table ─────────────────────────────────────────
        "All-Term Score Breakdown":      "All-Term Score Breakdown",
        "Category":                      "Category",
        "Weight":                        "Weight",
        "Released":                      "Released",
        "Pending":                       "Pending",
        "Not Yet Released":              "Not Yet Released",
        "Score / 100":                   "Score / 100",
        "Grade":                         "Grade",
        "Scores":                        "Scores",
        "Contribution":                  "Contribution",

        # ── English department score columns ──────────────────────────────────
        "Conduct":                       "Conduct",
        "Class Participation":           "Class Participation",
        "Homework & Assignments":        "Homework & Assignments",
        "Quiz":                          "Quiz",
        "Mid-Term Exam":                 "Mid-Term Exam",
        "Final Exam":                    "Final Exam",
        "Final Report":                  "Final Report",
        "Final Report Score":            "Final Report Score",

        # ── Chinese department score columns ──────────────────────────────────
        "Behavior":                      "Behavior",
        "Homework":                      "Homework",
        "Quizzes":                       "Quizzes",
        "Final Test":                    "Final Test",

        # ── Chinese grading scale ─────────────────────────────────────────────
        "Grading Scale":                 "Grading Scale",
        "Excellent":                     "Excellent",
        "Very Good":                     "Very Good",
        "Good":                          "Good",
        "Average":                       "Average",
        "Failure":                       "Failure",

        # ── Status / pass-fail labels ─────────────────────────────────────────
        "Status":                        "Status",
        "PASSED":                        "PASSED",
        "FAILED":                        "FAILED",
        "PASSING":                       "PASSING",
        "FAILING":                       "FAILING",
        "Passed":                        "Passed",
        "Failed":                        "Failed",
        "Incomplete":                    "Incomplete",

        # ── Year-to-date section ──────────────────────────────────────────────
        "Year-to-Date Average":          "Year-to-Date Average",
        "Cumulative Average":            "Cumulative Average",
        "OVERALL PASSING":               "OVERALL PASSING",
        "OVERALL FAILING":               "OVERALL FAILING",
        # pct is substituted at render time via %(pct)s
        "Passing threshold: %(pct)s%%":  "Passing threshold: %(pct)s%%",
        "No term data has been released yet. Check back after your first term results.":
            "No term data has been released yet. Check back after your first term results.",

        # ── Signature area ────────────────────────────────────────────────────
        "Class Teacher":                 "Class Teacher",
        "Principal":                     "Principal",
        "Parent / Guardian":             "Parent / Guardian",

        # ── Print / accreditation footer ──────────────────────────────────────
        "Print Report":                  "Print Report",
        "Accredited by":                 "Accredited by",
    },

    # ──────────────────────────────────────────────────────────────────────────
    # KHMER (ភាសាខ្មែរ)
    # ──────────────────────────────────────────────────────────────────────────
    "km": {
        # ── Navigation ────────────────────────────────────────────────────────
        "Departments":                   "កម្មវិធីសិក្សា",
        "Teacher Portal":                "សម្រាប់គ្រូ",
        "Admin":                         "រដ្ធបាល",
        "HOD Approvals":                 "ការអនុម័តប្រធានផ្នែក",
        "Teacher Logout":                "ចាកចេញ (គ្រូ)",
        "Admin Logout":                  "ចាកចេញ (អ្នកគ្រប់គ្រង)",
        "Logout":                        "ចាកចេញ",

        # ── Site-wide ─────────────────────────────────────────────────────────
        "Student Term Report Portal":    "ប្រព័ន្ធរបាយការណ៍លទ្ធផលសិស្ស",
        "All rights reserved.":          "រក្សាសិទ្ធគ្រប់យ៉ាង",

        # ── Landing page ──────────────────────────────────────────────────────
        "Welcome to the Student Report Portal":
            "សូមស្វាគមន៍មកកាន់គេហទំព័ររបាយការណ៍សិស្សានុសិស្ស",
        "Select your department below to view a term report card.":
            "សូមជ្រើសរើសកម្មវិធីសិក្សាដូចខាងក្រោម ដើម្បីមើលលទ្ធផលសិក្សា",
        "English Department":            "ថ្នាក់សិក្សាភាសាអង់គ្លេស",
        "Chinese Department":            "ថ្នាក់សិក្សាភាសាចិន",

        # ── Login ─────────────────────────────────────────────────────────────
        "Parent / Guardian Login":       "ការចូលជំពូករបស់មាតា-បិតា / អ្នកអាណាព្យាបាល",
        "Chinese Department Login":      "ការចូលជំពូកថ្នាក់ភាសាចិន",
        "Enter your child's full name and the password provided by the school.":
            "បំពេញឈ្មោះពេញរបស់កូន និងលេខសំងាត់ដែលបានផ្តល់ដោយសាលា",
        "Student Full Name":             "ឈ្មោះពេញរបស់សិស្ស",
        "Child's Class (to clarify)":    "ថ្នាក់របស់កូន (ដើម្បីបញ្ជាក់)",
        "— Select Class —":              "— ជ្រើសរើសថ្នាក់ —",
        "Multiple students share this name. Please select your child's class to continue.":
            "មានសិស្សច្រើននាក់ប្រើឈ្មោះនេះ។ សូមជ្រើសរើសថ្នាក់របស់កូន ដើម្បីបន្ត",
        "Multiple students share this name. Please select your child's class.":
            "មានសិស្សច្រើននាក់ប្រើឈ្មោះនេះ។ សូមជ្រើសរើសថ្នាក់របស់កូន",
        "Parent Password":               "លេខសំងាត់មាតា-បិតា",
        "Enter password":                "វាយបញ្ចូលលេខសំងាត់",
        "View Report Card":              "មើលក្រដាសលទ្ធផល",
        "Wrong department?":             "ខុសថ្នាក់សិក្សាមែនទែ?",

        # ── Report card — school header ───────────────────────────────────────
        "Official Term Report Card":     "ក្រដាសលទ្ធផលសិក្សាផ្លូវការ",
        "Academic Year 2024 \u2013 2025":
            "ឆ្នាំសិក្សា ២០២៤ \u2013 ២០២៥",
        "Chinese Department \u2013 Official Term Report Card":
            "ថ្នាក់សិក្សាភាសាចិន \u2013 ក្រដាសលទ្ធផលសិក្សាផ្លូវការ",
        "CN Program \u2022 Academic Year 2024 \u2013 2025":
            "កម្មវិធី CN \u2022 ឆ្នាំសិក្សា ២០២៤ \u2013 ២០២៥",

        # ── Report card — student meta block ─────────────────────────────────
        "Student Name":                  "ឈ្មោះសិស្ស",
        "Student ID":                    "អត្តលេខសិស្ស",
        "No. in List":                   "លេខក្នុងបញ្ជី",
        "Class":                         "កម្រិតថ្នាក់/វេនសិក្សា",
        "Department":                    "ថ្នាក់សិក្សា",
        "Chinese (CN Program)":          "ភាសាចិន (កម្មវិធី CN)",
        "Report Generated":              "កាលបរិច្ឆេទ",

        # ── Report card — term selector ───────────────────────────────────────
        "View Term":                     "ជ្រេីសរើសវគ្គសិក្សា",
        "Master Overview (All Terms)":   "ទិដ្ឋភាពសរុប (គ្រប់ក្រតា)",
        "Term %(n)s":                    "ក្រតាទី %(n)s",
        "Term %(n)s Report":             "របាយការណ៍ក្រតាទី %(n)s",
        "Term %(n)s Report \u2013 Chinese Department":
            "របាយការណ៍ក្រតាទី %(n)s \u2013 ថ្នាក់សិក្សាភាសាចិន",

        # ── Report card — score table ─────────────────────────────────────────
        "All-Term Score Breakdown":      "សង្ខេបពិន្ទុគ្រប់ក្រតា",
        "Category":                      "រាយមុខវិជ្ជា",
        "Weight":                        "ពិន្ទុសរុប",
        "Released":                      "មានរបាយការណ៍",
        "Pending":                       "មិនទាន់មានរបាយការណ៍",
        "Not Yet Released":              "មិនទាន់ចេញផ្សាយ",
        "Score / 100":                   "ពិន្ទុ / ១០០",
        "Grade":                         "និទ្ទេស",
        "Scores":                        "ពិន្ទុ",
        "Contribution":                  "ការចូលរួម",

        # ── English department score columns ──────────────────────────────────
        "Conduct":                       "ឥរិយាបថ",
        "Class Participation":           "ការចូលរួមក្នុងថ្នាក់",
        "Homework & Assignments":        "កិច្ចការផ្ទះ និងកិច្ចការក្នុងថ្នាក់",
        "Quiz":                          "កិច្ចការសាលា",
        "Mid-Term Exam":                 "ការប្រឡងពាក់កណ្ដាលវគ្គសិក្សា",
        "Final Exam":                    "ការប្រឡងបញ្ចប់វគ្គសិក្សា",
        "Final Report":                  "លទ្ធផលសរុប",
        "Final Report Score":            "ពិន្ទុលទ្ធផលសរុប",

        # ── Chinese department score columns ──────────────────────────────────
        "Behavior":                      "ឥរិយាបថ",
        "Homework":                      "កិច្ចការផ្ទះ",
        "Quizzes":                       "កិច្ចការសាលា",
        "Final Test":                    "ការប្រឡងបញ្ចប់វគ្គសិក្សា",

        # ── Chinese grading scale ─────────────────────────────────────────────
        "Grading Scale":                 "មាត្រដ្ឋានការវាយតម្លៃ",
        "Excellent":                     "ល្អឆ្នើម",
        "Very Good":                     "ល្អខ្លាំង",
        "Good":                          "ល្អ",
        "Average":                       "មធ្យម",
        "Failure":                       "ធ្លាក់",

        # ── Status / pass-fail labels ─────────────────────────────────────────
        "Status":                        "ស្ថានភាព",
        "PASSED":                        "ជាប់",
        "FAILED":                        "ធ្លាក់",
        "PASSING":                       "កំពុងជាប់",
        "FAILING":                       "កំពុងធ្លាក់",
        "Passed":                        "ជាប់",
        "Failed":                        "ធ្លាក់",
        "Incomplete":                    "មិនពេញ",

        # ── Year-to-date section ──────────────────────────────────────────────
        "Year-to-Date Average":          "មធ្យមភាគគិតត្រឹមបច្ចុប្បន្ន",
        "Cumulative Average":            "មធ្យមភាគសរុប",
        "OVERALL PASSING":               "ជាប់ទូទៅ",
        "OVERALL FAILING":               "ធ្លាក់ទូទៅ",
        # %(pct)s is substituted at render time; %% becomes a literal %
        "Passing threshold: %(pct)s%%":  "ជម្រើសសិទ្ធ: %(pct)s%%",
        "No term data has been released yet. Check back after your first term results.":
            "មិនទាន់មានទិន្នន័យក្រតាណាមួយត្រូវបានចេញផ្សាយទេ។ "
            "សូមត្រឡប់មកវិញបន្ទាប់ពីលទ្ធផលក្រតាទី ១",

        # ── Signature area ────────────────────────────────────────────────────
        "Class Teacher":                 "គ្រូប្រចាំថ្នាក់",
        "Principal":                     "នាយក",
        "Parent / Guardian":             "មាតា-បិតា / អ្នកអាណាព្យាបាល",

        # ── Print / accreditation footer ──────────────────────────────────────
        "Print Report":                  "បោះពុម្ពរបាយការណ៍",
        "Accredited by":                 "ទទួលស្គាល់ដោយ",
    },
}
