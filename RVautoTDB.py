
import tkinter as tk
from tkinter import messagebox, filedialog
import os
import subprocess


root = tk.Tk()
root.title("Garmin-inställningar & instruktioner")
# Hämta skärmens bredd och höjd
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
# Fönsterbredd
window_width = 500
# Fönsterhöjd = hela skärmen
window_height = screen_height
# Placera längst till höger, från toppen
x_pos = screen_width - window_width
y_pos = 0
root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

instructions = (
    "Välkommen till RVautoTDB\n"
    "Koden är framtagen av Rasmus Viertel för att förenkla Träningsdagboken 20021018 av Mats Troeng\n\n"
    "Så här använder du scriptet för Garmin Connect och Excel:\n\n"
    "1. Installera Python om du inte redan har det.\n"
    "2. Fyll i dina Garmin Connect-uppgifter och Excel-fil i inställningar.txt som bör ligga i samma mapp.\n"
    "3. Se till att din Excel-makrofil (t.ex. 2024-2025 dagbok.xlsm) ligger i samma mapp.\n"
    "4. Skapa en virtuell miljö i mappen:\n"
    "     1. Skriv i terminalen: python -m venv .venv\n"
    "     2. Aktivera miljön: .\\.venv\\Scripts\\activate\n"
    "   - Installera sedan paketen i miljön: pip install pandas openpyxl garminconnect\n"
    "5. Kör scriptet: python import_garmin_to_excel.py\n"
    "6. Scriptet hämtar träningsdata från Garmin Connect och fyller i din Excel-makrofil.\n\n"
    "Om något saknas eller är fel i inställningar.txt får du ett tydligt felmeddelande.\n"
    "Om din Dagbok inte slutar på .xlsm kontakta rasmus@viertel.se\n"
    "Läs mer om kodens funktioner på https://rasmus.viertel.se/\n"
)

# Dela upp instruktionerna i två delar för att placera knappen rätt
instructions_top_1 = (
    "Välkommen till RVautoTDB\n"
    "Koden är framtagen av Rasmus Viertel för att förenkla Träningsdagboken 20021018 av Mats Troeng\n\n"
    "Så här använder du scriptet för Garmin Connect och Excel:\n\n"
    "1. Installera Python om du inte redan har det.\n"
)
instructions_top_2 = (
    "2. Fyll i dina Garmin Connect-uppgifter och Excel-fil i inställningar.txt som bör ligga i samma mapp.\n"
    "3. Se till att din Excel-makrofil (t.ex. 2024-2025 dagbok.xlsm) ligger i samma mapp.\n"
    "4. Skapa en virtuell miljö i mappen:\n"
)
instructions_bottom = (
    "Kom ihåg att fylla i inställningar.txt med dina uppgifter. Annars kommer scriptet att krascha.\n"
    "Om din Dagbok inte slutar på .xlsm kontakta rasmus@viertel.se\n"
    "Läs mer om kodens funktioner på https://rasmus.viertel.se/\n"
)

# Visa första delen av instruktionerna
lbl_info_top_1 = tk.Label(root, text=instructions_top_1, font=("Arial", 10), justify="left", wraplength=480)
lbl_info_top_1.pack(pady=10)

# Knapp för att ladda ner Python
def open_python_download():
    import webbrowser
    webbrowser.open("https://www.python.org/downloads/")

btn_download_python = tk.Button(root, text="Ladda ner här", command=open_python_download, font=("Arial", 10), width=18)
btn_download_python.pack(pady=5)

# Visa andra delen av instruktionerna
lbl_info_top_2 = tk.Label(root, text=instructions_top_2, font=("Arial", 10), justify="left", wraplength=480)
lbl_info_top_2.pack(pady=0)

# Knapp för att ladda ner Python


# Knapp för att öppna PowerShell
def open_powershell():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    subprocess.Popen(["powershell", "-NoExit"], cwd=script_dir)

btn_open_ps = tk.Button(root, text="Öppna PowerShell", command=open_powershell, font=("Arial", 10), width=18)

btn_open_ps.pack(pady=5)

# Instruktionstext mellan knapp och kommandoruta
lbl_venv_info = tk.Label(root, text="Kör detta kommando vid första användning:", font=("Arial", 10), justify="left")
lbl_venv_info.pack(pady=(10,2))



# Visa 'python -m venv .venv' som kopierbar kodrad (Entry-widget)
venv_cmd_entry = tk.Entry(root, font=("Consolas", 11), fg="blue", width=30, justify="left")
venv_cmd_entry.insert(0, "python -m venv .venv")
venv_cmd_entry.config(state="readonly")


venv_cmd_entry.pack(pady=2)

# Instruktion mellan kodsnutterna
lbl_activate_info = tk.Label(root, text="Detta kommando ska du köra varje gång:", font=("Arial", 10), justify="left")
lbl_activate_info.pack(pady=(10,2))

# Visa '.\\.venv\\Scripts\\activate' som kopierbar kodrad (Entry-widget)
activate_cmd_entry = tk.Entry(root, font=("Consolas", 11), fg="blue", width=30, justify="left")
activate_cmd_entry.insert(0, ".\\.venv\\Scripts\\activate")
activate_cmd_entry.config(state="readonly")
activate_cmd_entry.pack(pady=2)

# Instruktionstext för pip install
lbl_pip_info = tk.Label(root, text="Kor detta komando vid första användning:", font=("Arial", 10), justify="left")
lbl_pip_info.pack(pady=(10,2))
# Kopierbar kodrad för pip install
pip_cmd_entry = tk.Entry(root, font=("Consolas", 11), fg="blue", width=40, justify="left")
pip_cmd_entry.insert(0, "pip install pandas openpyxl garminconnect")
pip_cmd_entry.config(state="readonly")
pip_cmd_entry.pack(pady=2)

# Instruktionstext för att köra scriptet
lbl_run_info = tk.Label(root, text="Kör detta kommandot för att fylla i träningsdagboken:", font=("Arial", 10), justify="left")
lbl_run_info.pack(pady=(10,2))
# Kopierbar kodrad för att köra scriptet
run_cmd_entry = tk.Entry(root, font=("Consolas", 11), fg="blue", width=40, justify="left")
run_cmd_entry.insert(0, "python import_garmin_to_excel.py")
run_cmd_entry.config(state="readonly")
run_cmd_entry.pack(pady=2)

lbl_info_bottom = tk.Label(root, text=instructions_bottom.replace("     1. Skriv i terminalen: python -m venv .venv\n", ""), font=("Arial", 10), justify="left", wraplength=480)
lbl_info_bottom.pack(pady=0)


# Funktion för att välja Python-fil och visa filnamnet i instruktionen

root.mainloop()
