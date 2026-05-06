"""
init_project.py
---------------
Run this script ONCE to bootstrap the project.
It creates the required folder structure and a starter grades.xlsx file
with five sheets: Students, Grades, Teachers, Admins, and ApprovalStatus.

SCHEMA OVERVIEW
───────────────
• Students sheet:
    StudentID | Name | ClassLabel | ParentPassword

    CRITICAL: Parents log in using Student Full Name + ParentPassword
    (not StudentID). See PARENT_LOGIN_REFACTORING.md for details.

• Grades sheet:
    StudentID | Term | Conduct | CP | HW_ASS | QUIZ | MidTerm | Final | FinalReport

• Teachers sheet:
    Username | Password | Role (HOD or Teacher)

• Admins sheet:
    Username | Password

• ApprovalStatus sheet:
    StudentID | Term | Approved | RequestNote
    (Starts empty; populated as scores are approved by HOD)

IMPORTANT: If you are upgrading from a previous version, delete the old
grades.xlsx first — the schema has evolved significantly.

POST-INITIALIZATION STEPS
──────────────────────────
1. Run app.py: python app.py
2. If ParentPassword cells are blank, run: python fill_missing_passwords.py
3. Distribute ParentPasswords to parents securely (not via email/SMS)

Required libraries:
    pip install pandas openpyxl flask
"""

import os
import pandas as pd

# ── 1. Create folders ──────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

folders = [
    os.path.join(BASE_DIR, "templates"),
    os.path.join(BASE_DIR, "static"),
]

for folder in folders:
    os.makedirs(folder, exist_ok=True)
    print(f"[OK] Folder ready: {folder}")

# ── 2. Create grades.xlsx (five sheets) ───────────────────────────────────────
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")

# ── Sheet 1: Students ─────────────────────────────────────────────────────────
# One row per student. Parents log in using the Name + ParentPassword columns.
# StudentID is used internally for grade lookups; ClassLabel handles parens like L6T2(2).
students_data = {
    "StudentID": [
        "D01965","D00517","D00700","D01836","D01381","D01520","D01146","D01740","D02229","D01328","D00511","D01514",
        "D01905","D02202","D02125","D01791","D02192","D01123","D01367","D01333","D02145","D01757","D01507","D00583",
        "D01496","D01754","D02053","D01738","D01852","D01796","D02216","D00334","D02180","D02126","D01216","D01534",
        "D01774","D01782","D00983","D01697","D01778","D01739","D02314","D02123","D02227","D00651","D02235","D00735",
        "D01783","D01404","D01786","D02143","D01646","D01225","D02248","D01124","D00952","D01510","D02132","D02073",
        "D02234","D01851","D00390","D02205","D02157","D01509","D01807","D02237","D00784","D02179","D00612","D02108",
        "D00675","D01528","D00347","D02058","D00241","D01550","D00311","D00330","D00787","D01990","D01845","D00138",
        "D00981","D00201","D01382","D00328","D00147","D01806","D00515","D00232","D00867","D01779","D01687","D01565",
        "D01948","D01841","D02067","D01044","D01964","D00307","D01531","D00883","D01938","D01860","D01542","D02144",
        "D02070","D02129","D01775","D02284","D01318","","","","","D00385","D01279","D00176",
        "D01784","D01431","D00175","D02075","D00249","D01532","D00592","D00247","D01297","D00323","D00876","D00380",
        "D00075","D00155","D00180","D00121","D01284","D00875","D00253","D00062","D01166","D00150","D01035","D00428",
        "D01777","D00196","","D00194","D01745","D00690","D01536","D00112","D01615","D00035","D00036","D00577",
        "D00110","D01785","D01787","D00292","D01057","D00140","D01358","D02249","","D01434","D02173","D00128",
        "D00384","D00037","D00039","D00043","D00049","D02185","D00064","D02196","D00072","D00174","D02263","D02257",
        "D01377","D02254","D02270","D02264","D02265","D02267","D02256","D02273","D02247","D02255","D00705","D01092",
        "D02342","D02338","D02337","D02336","D02370","","","D02373","D01000","D02259","D02178","D01925",
        "D02177","D02261","D01933","D02258","D02278","D02260","D02245","D02262","D02266","D02269","D02268","D02302",
        "D02319","D01068","","D01244","D00604","D01626","D01902","D00154","D01471","D01728","D00243","D00377",
        "","","D01729","D00749","D01645","D01853","D01571","D01200","D00267","D01109","D01254","D02293",
        "D00429","D00831","D00946","D00880","D01847","D00086","D00287","D01620","D00684","D01356","D00346","D01402",
        "D01384","E00191","D02244","D00885","D00399","D00012","D00034","D00053","D01088","D01529","D00097","D00096",
        "D01465","","","","","","","","","","","",
        "","","","","","","","","","D02061","D02189","D02015",
        "D00990","D00123","D01751","D01727","D02297","D02344","D00089","D00045","D00061","D00066","D00105","D00091",
        "D00090","D00306","D00042","D00046","D00126","","D01833","D01549","D01951","D00606","D01859","D00972",
        "D02047","D02044","D01618","D01302","D01943","D00767","D01482","D02165","D01042","D01809","D02148","D01206",
        "D01187","D01309","D01336","D01335","D00605","D01115","D01181","D00886","D01437","D01867","D01903","D01953",
        "D02285","D01036","D01477","D00424","D01084","D01334","D00272","D01108","D00447","D00659","D01957","D01674",
        "D00570","D01513","D01183","D01559","D01136","D01749","D01199","D00598","D00149","D01430","D01119","D01394",
        "D00693","D01293","D00595","D00409","D01846","D00915","D01445","D01628","D01548","D00611","D00610","D01252",
        "D00921","D01275","D00679","D00955","D00300","D01233","D01305","D01398","D00363","D00257","D00254","D00701",
        "D00289","D01208","D00628","D00863","D00228","D00153","D00852","D00173","D00480",
    ],
    "Name": [
        "Chhuor Chunminh","HuyChhun Yuly","Khim Sufong","Mony Chhayly","Nan Keo kosomak","Phon VengKheang","Samrith Phoumen","Thy Chanreach","Thy Ly Hour","Vann Gechly","Hong Rithiponleu","Limleng Boromsastra",
        "Ann Chanbormey","Mo Pok Ying","Chhunsarith Monypich","Ha Haihong","Horn Pechrithnak","Kuoch Chhorvivan","Lor Khongminh","Mony Chansomonea","Sokna Heng Heng","Theam Vatana","Thoeun Bunnareach","Vicheth Seakney",
        "Voeung Lyherin","Soun Navin","Brak Channaly","Chhoeun Sarun Oudom","Dy Sreysros","Lail Bunhok","Ny Sophanphilin","Pheak Phearum","Pich Ly Eng","Rin Ratanakpanha","Sa Em Panhavoan","Sak Mengyu",
        "San Rayut","Sey Samerdy","Tang Kanghor","Um Chhaylenghong","Ven Vireakssatya","Vorn Pichjulyka","Von Chanto","Monyreak Sokhorng","Beun Ponleu","Chea Chiengsou","Chuot Baravin","Mel Duong Khae",
        "Oung Sosakreach","Pheng Ratanak","Ry Lyching","Sokna Sophanika","Sokun Youming","Tang Chousor","Teng Lyhorng","Theung Chanmonika","Vann Seuchhay","Hoeng Mengheng","Chen kaknika","Chhay Rithyvireak",
        "Din Soveacha","Dy Minea","Eim Kimly","Heat Daneth","Hem Noudy","Ly Kimyoubin","Mao Chhunheng","Pachek Vireak","Phally GiGi","Pich Yu Eii","Ravut Ratanaksambat","Run Oudom",
        "Sa Em Molyvann","Thai Soksreyneang","Ung Bunn Xiang","Ung Kimmey","Vut Nachael","Yang Narith","Luy Ratanakjason","Moeuth Menghong","Tech Lymeng","Ni Horng","Chea Sourajel","Hor Chealong",
        "Im Vorada","Lim Ratana","Nan Sithasak","Sam Ratanakpitou","Sao Sopheary","Seam Lychhean","Sim Kimsan","Sok Leangmeng","Sokna Sereybrosethreaksmey","Tech Ariya","Thorng SiveIng","Touch Vary",
        "Aek Theara","Kheng Bunleap","Khorn Dara Reach","Leanghort Seng Im","Ly Dona","Ly Seav Ing","Nhim Souphorn","Ou Pengleang","Oung Senghour","Samout Kakda","Seong Menghorng","Sokna Sophanita",
        "Song Somalortey","Thorng Mondolkeo","Van Neth","Yin limhak","Lounh Hengleap","Sok Lyly","Soeun Kanhanika","Chhean ChhengVan","Leng Meyling","Heam Sokheng","Pheng Kimhorng","Ratha Jingping",
        "Ry Rany","Say Chhunheng","Sean Tonghai","Seang Ousaphea","Sy Usinh","Theung Chansomonor","Touch Kimlang","Vin Julida","Em Chanmolika","Chum Yoshiko Mori","Khe David","Lim Sivkea",
        "Luy Layinn","Neak Siv Ing","Nguon Chu In","Rithy Monita","San Huy Sing","Sat Kimly","Say Chhengly","Seth Sivatey","So Hokyean","Tang Thai Eang","Ty KimLeang","Vann Seulong",
        "Ven Vortey","Yuth Chhenghun","Chhel Socheata","Meng Heanchinna","Tes Lyhor","Thet Sreyka","Vary Pharaket","Heng Singher","Hour Mengkong","Mut Chetvireak","Neak Menghour","Pov Sombathvirakyu",
        "Ratha Bunleang","Ry Liya","Soem Sreypich","Sokha Nara","Tang Selina","Tang Vuochlang","Voeun Kuymeng","Sen Mouychheng","Seat Chheangvorn","Tenghab Bunma","Heng Kanika","Kong Virak",
        "Koy Lyhor","Nem Neakvireak","Nguon Ratanakrobin","Ou Yeak Jinh","Phen Oudom","Sai Kimhab","Sean Seang Hai","Soeun Naiey","Srun Theathahena","Chea Soheng","An Bopha","Chan Ratana",
        "Houn Vanna","Kong Socheata","Moy Limith","Nhuk Sreymean","Nhuk Sreymeas","Nuk Hoching","Phat Oudom","Samnang Sonnra Vitou","San Rithmany","Yon Narin","Pen Chanvireak","Chhon Hanny",
        "Kae Sreyleak","Naev Vantha","Naev Kabong","Pheakdey Sreylin","Phi Sokhai","Cheut Thavisiny","Veasana Vayut","Sery Samrith","Chev Chumying","Chheo Athicheth","Lay Mengly","Ly Simheng",
        "Meach Nita","Min Rot","Nypich Sopheakstra","Raksmey Dara","Ry Macheng","Samnang Sovotey","Samoeung Lyza","Thoeun Chan Bormey","Thy Seav Im","Yon Darong","Yon Reaksa","Chan Chanthy",
        "Chan Makara","Chin Panha","Phea Sophin","Ann Borith","Chhorn Pisey","Heng Pouthirech","Kham Senghok","Luy Raksmey Rithy Sky","Ly Channary","Ly Kimmandy","Ly Vouchnea","Oeung Chheang E",
        "Ny Sreytom","Leng Sreykeo","Sanith Monyratanak","Run Layheang","Theam Thavora","San Suothanorn","Nhem Cheysithipol","Ly Mara","Yim Sarun Ouchin","Naet Panharith","Kham Ratanakpiseth","Sok Leang Heng",
        "Vanna Sovanpanha","Va Pisal","Vy Masurky","Choeun Davina","Heng Sive Pao","Kuoch Hankla","Meas Chanliza","Nov Fuminh","Pen Marida","Pheak Bunlong","Rith Jesda","Theam Kanika",
        "Voeung Limhorng","Wu Xiaomei","Em Vichea","Som Sovanntha","Hak Sodalin","Hem Menghorng","Mok Chamrong","Rithy Pisith","Say Sokpiseth","Thon Mary","Tun Vannisa","Tun Vatana",
        "Nov Maneth","Thun Sothich","Theang Bunsreyka","Khon Samnang","Than Sreynet","Net sreynich","Prak Siv Eing","Neth Viraeboth","Sithi Malihor","Cheang Sovanvireak","Kavan Sok","Pleav Kanek",
        "Kateng Sathea","Mouk Kapinh","Katta Bronh","Tangki Pouth","Klem Davy","Soeun Sreynich","Thy Sreyleap","Cheang Khnach","Mey Srey vit","Bo Sokleap","Boel Sopheanu","Bun Sopanhareach",
        "Ly Chhunlong","Pheng Lysivming","Sreng Hengkim","Sreng Ve ha","San Ratana","Phornphan Nophea","Leanghort Seng Er","Ourn Maniratanak","Seth Houjing","Sok Leangkim","Hun Bunlong","Leang Ing Ing",
        "Leanghort Senghok","Ly Seav Er","Ong Byle","Ourn Ousa","Song Ratanakchenfong","Seat Niza","Chong Jing In","Chong Siv Tieng","Heang Siv Chhin","Heng Menghouy","Hom Vitou","Keat Sokunpitou",
        "Leak Vouchleang","Leng Kimlang","Leng Narin","Lugowski Julia","Ly Seuseu","Ly Sing","Phat Panhchak Champey","Pheab Yu fim","Rin Sokuntheapanha","Thak JV","Thea Gihor","Theam Meyling",
        "Samit Kalika","Sung Livhor","Hak Lymeng","Hak Solika","Heng Mengly","Kheoun Nangsiunna","Ly Panhavichetra","Ry Chumchang","Sam Ratanaksatya","Say Sorida","Sela Yagkur","Soeun KimMengNgoun",
        "Sameth Chhun E","Ty Sonlik","Chhom Lyly","Hor Mouyly","Koy Nikor","Kung Sereyvathna","Ly Kim Gech Heng","Naet Vireakyuk","Nara Maraneth","Nimsareurn Munita","Ny Lyhor","Ourn Monyreach",
        "Phat Meyling","Ra Rabe","Saroeurn Davika","Theam Meng Eang","Tola Saychhom","Try Sovanpanha","Ann Lyhour","Aum Silim","Chhoeun Thammai","Eart Amara","Hong Ly Ing","Hout Monny",
        "Lugowski Maya","Ly Tedana","Ratha Jing Er","Seng Sopanha","So Seanghoy","Tenghab Leakhena","Yun Jiyin","Chea Tana","Chong Vengxing","Heang Seavlay","Heang Seavlong","Kham Ratanakkuntheany",
        "Kong Sokpiseth","Ly Chouseang","Ly Mei Mei","Oung Erithyraksmey","Phat Roat chhouk","Theam Meychou","Tho Li Eng","Va Sakana","Bour Panhareach","Choeurn Vichetmongkol","Chou Sotornta","Hak Suchhing",
        "Hong Lyhuor","Hour Meylim","Kong Chansreypich","Lim Tongheang","Meng Heankimtav","Ong Sindy","Pheng Mengkheang","Srun Manath","Vuthy Bayling",
    ],
    "ClassLabel": [
        "L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3",
        "L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)","L1T3(2)",
        "L1T3(2)","L1T3(2)","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3",
        "L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T5","L2T5","L2T5","L2T5",
        "L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L3T4","L3T4",
        "L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4",
        "L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L3T4","L5T2","L5T2",
        "L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2",
        "L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2",
        "L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L6T3","L6T3","L6T3",
        "L6T3","L6T3","L6T3","L6T3","L6T3","L6T3","L6T3","L6T3","L6T3","L7T1","L7T1","L7T1",
        "L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1","L7T1",
        "L7T1","L7T1","L7T1","L8T3","L8T3","L8T3","L8T3","L9T2","L9T2","L9T2","L9T2","L9T2",
        "L9T2","L9T2","L9T2","L9T2","L9T2","L9T2","L9T2","L9T2","L9T2","L9T2","L11T4","L11T4",
        "L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L1T1","L1T1",
        "L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1",
        "L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T1","L1T2","L1T2","L1T2",
        "L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2","L1T2",
        "L1T2","L1T2","L1T2","L6T2","L6T2","L6T2","L6T2","L6T2","L6T2","L6T2","L6T2","L6T2",
        "L6T2","L6T2","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)","L6T2(2)",
        "L6T2(2)","L6T2(2)","L6T2(2)","L9T4","L9T4","L9T4","L9T4","L9T4","L9T4","L9T4","L9T4","L9T4",
        "L9T4","L9T4","L9T4","L9T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4","L11T4",
        "L11T4","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)",
        "L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L1T1(2)","L4T3","L4T3","L4T3",
        "L4T3","L4T3","L4T3","L4T3","L4T3","L4T3","L10T4","L10T4","L10T4","L10T4","IELTS","IELTS",
        "IELTS","IELTS","IELTS","IELTS","IELTS","IELTS","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3",
        "L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3","L1T3",
        "L1T3","L1T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3","L2T3",
        "L2T3","L2T3","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L2T5",
        "L2T5","L2T5","L2T5","L2T5","L2T5","L2T5","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2",
        "L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L4T2","L5T2","L5T2","L5T2","L5T2","L5T2",
        "L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L5T2","L6T1","L6T1","L6T1","L6T1",
        "L6T1","L6T1","L6T1","L6T1","L6T1","L6T1","L6T1","L6T1","L6T1",
    ],
    "ParentPassword": [
        "D01965","D00517","D00700","D01836","D01381","D01520","D01146","D01740","D02229","D01328","D00511","D01514",
        "D01905","D02202","D02125","D01791","D02192","D01123","D01367","D01333","D02145","D01757","D01507","D00583",
        "D01496","D01754","D02053","D01738","D01852","D01796","D02216","D00334","D02180","D02126","D01216","D01534",
        "D01774","D01782","D00983","D01697","D01778","D01739","D02314","D02123","D02227","D00651","D02235","D00735",
        "D01783","D01404","D01786","D02143","D01646","D01225","D02248","D01124","D00952","D01510","D02132","D02073",
        "D02234","D01851","D00390","D02205","D02157","D01509","D01807","D02237","D00784","D02179","D00612","D02108",
        "D00675","D01528","D00347","D02058","D00241","D01550","D00311","D00330","D00787","D01990","D01845","D00138",
        "D00981","D00201","D01382","D00328","D00147","D01806","D00515","D00232","D00867","D01779","D01687","D01565",
        "D01948","D01841","D02067","D01044","D01964","D00307","D01531","D00883","D01938","D01860","D01542","D02144",
        "D02070","D02129","D01775","D02284","D01318","847293","625974","512847","739156","D00385","D01279","D00176",
        "D01784","D01431","D00175","D02075","D00249","D01532","D00592","D00247","D01297","D00323","D00876","D00380",
        "D00075","D00155","D00180","D00121","D01284","D00875","D00253","D00062","D01166","D00150","D01035","D00428",
        "D01777","D00196","481629","D00194","D01745","D00690","D01536","D00112","D01615","D00035","D00036","D00577",
        "D00110","D01785","D01787","D00292","D01057","D00140","D01358","D02249","693215","D01434","D02173","D00128",
        "D00384","D00037","D00039","D00043","D00049","D02185","D00064","D02196","D00072","D00174","D02263","D02257",
        "D01377","D02254","D02270","D02264","D02265","D02267","D02256","D02273","D02247","D02255","D00705","D01092",
        "D02342","D02338","D02337","D02336","D02370","258741","876543","D02373","D01000","D02259","D02178","D01925",
        "D02177","D02261","D01933","D02258","D02278","D02260","D02245","D02262","D02266","D02269","D02268","D02302",
        "D02319","D01068","341987","D01244","D00604","D01626","D01902","D00154","D01471","D01728","D00243","D00377",
        "564821","729456","D01729","D00749","D01645","D01853","D01571","D01200","D00267","D01109","D01254","D02293",
        "D00429","D00831","D00946","D00880","D01847","D00086","D00287","D01620","D00684","D01356","D00346","D01402",
        "D01384","E00191","D02244","D00885","D00399","D00012","D00034","D00053","D01088","D01529","D00097","D00096",
        "D01465","158346","947283","376892","645271","819364","297845","532768","418765","761234","893451","624815",
        "375964","582193","641728","769512","453698","286574","714936","897542","325687","D02061","D02189","D02015",
        "D00990","D00123","D01751","D01727","D02297","D02344","D00089","D00045","D00061","D00066","D00105","D00091",
        "D00090","D00306","D00042","D00046","D00126","956238","D01833","D01549","D01951","D00606","D01859","D00972",
        "D02047","D02044","D01618","D01302","D01943","D00767","D01482","D02165","D01042","D01809","D02148","D01206",
        "D01187","D01309","D01336","D01335","D00605","D01115","D01181","D00886","D01437","D01867","D01903","D01953",
        "D02285","D01036","D01477","D00424","D01084","D01334","D00272","D01108","D00447","D00659","D01957","D01674",
        "D00570","D01513","D01183","D01559","D01136","D01749","D01199","D00598","D00149","D01430","D01119","D01394",
        "D00693","D01293","D00595","D00409","D01846","D00915","D01445","D01628","D01548","D00611","D00610","D01252",
        "D00921","D01275","D00679","D00955","D00300","D01233","D01305","D01398","D00363","D00257","D00254","D00701",
        "D00289","D01208","D00628","D00863","D00228","D00153","D00852","D00173","D00480",
    ],
}

students_df = pd.DataFrame(students_data, columns=["StudentID", "Name", "ClassLabel", "ParentPassword"])

# ── Sheet 2: Grades ───────────────────────────────────────────────────────────
# One row per (StudentID, Term) — up to 4 rows per student.
# Name and ParentPassword are NOT duplicated here; looked up from Students sheet.
# Weighted formula: Conduct*0.05 + CP*0.05 + HW_ASS*0.15 + QUIZ*0.15 + MidTerm*0.25 + Final*0.35
grades_data = {
    "StudentID":   ["S001",  "S001",  "S002",  "S003"],
    "Term":        [1,       2,       1,       1],
    "Conduct":     [88.0,    90.0,    78.0,    87.0],
    "CP":          [90.0,    88.0,    80.0,    85.0],
    "HW_ASS":      [85.0,    87.0,    72.0,    92.0],
    "QUIZ":        [80.0,    83.0,    74.0,    89.0],
    "MidTerm":     [78.0,    82.0,    70.0,    88.0],
    "Final":       [82.0,    85.0,    75.0,    90.0],
    "FinalReport": [81.85,   84.65,   73.55,   89.25],
    # S001-T1: (88*.05)+(90*.05)+(85*.15)+(80*.15)+(78*.25)+(82*.35) = 81.85
    # S001-T2: (90*.05)+(88*.05)+(87*.15)+(83*.15)+(82*.25)+(85*.35) = 84.65
    # S002-T1: (78*.05)+(80*.05)+(72*.15)+(74*.15)+(70*.25)+(75*.35) = 73.55
    # S003-T1: (87*.05)+(85*.05)+(92*.15)+(89*.15)+(88*.25)+(90*.35) = 89.25
}

grades_df = pd.DataFrame(
    grades_data,
    columns=["StudentID", "Term", "Conduct", "CP", "HW_ASS",
             "QUIZ", "MidTerm", "Final", "FinalReport"],
)

# ── Sheet 3: Teachers ─────────────────────────────────────────────────────────
# One row per teacher account. Role can be "Teacher" or "HOD".
# Add more rows here or directly in Excel.
teachers_data = {
    "Username": ["Teacher Red",   "Teacher Jenne", "Teacher Ruby", "Teacher Ann", "Teacher Neth", "Teacher Vitou", "Teacher Norak", "Teacher Chamroeun"],
    "Password": ["P@ss4Eng26", "ReadWrit#9", "Term4Logic", "AlphaBeta!7", "Grammar88*", "OxfordPct$2", "Literature#1", "SyntaxError5"],
    "Role":     ["HOD",         "Teacher",       "Teacher",       "Teacher",     "Teacher",       "Teacher",       "Teacher",       "Teacher"],
}

teachers_df = pd.DataFrame(teachers_data, columns=["Username", "Password", "Role"])

# ── Sheet 4: Admins ────────────────────────────────────────────────────────────
# Student Affairs admin accounts. Add more rows as needed.
admins_data = {
    "Username": ["admin"],
    "Password": ["Admin@TEG2026"],
}

admins_df = pd.DataFrame(admins_data, columns=["Username", "Password"])

# ── Sheet 5: ApprovalStatus ───────────────────────────────────────────────────
# Starts empty — all class+term combos hidden from parents until approved.
approval_df = pd.DataFrame(columns=["ClassLabel", "Term", "Approved"])

if os.path.exists(EXCEL_PATH):
    print(f"[SKIP] grades.xlsx already exists — not overwriting.")
    print(f"       Delete it and re-run this script to apply the new schema.")
else:
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        students_df.to_excel(writer,  sheet_name="Students",       index=False)
        grades_df.to_excel(writer,    sheet_name="Grades",         index=False)
        teachers_df.to_excel(writer,  sheet_name="Teachers",       index=False)
        admins_df.to_excel(writer,    sheet_name="Admins",         index=False)
        approval_df.to_excel(writer,  sheet_name="ApprovalStatus", index=False)
    print(f"[OK] Created: {EXCEL_PATH}")
    print(f"     Sheet 'Students':       {len(students_df)} students (parents login with Name + ParentPassword)")
    print(f"     Sheet 'Grades':         sample grade records")
    print(f"     Sheet 'Teachers':       {len(teachers_df)} teacher accounts")
    print(f"     Sheet 'Admins':         {len(admins_df)} admin account(s) — admin / Admin@TEG2026")
    print(f"     Sheet 'ApprovalStatus': empty (updated as HOD approves scores)")
    print()
    print("NEXT STEPS:")
    print("  1. python fill_missing_passwords.py  (fill any blank ParentPassword cells)")
    print("  2. python app.py                     (start the Flask application)")
    print("  3. Visit http://localhost:5000/login (test parent login with Name + ParentPassword)")
