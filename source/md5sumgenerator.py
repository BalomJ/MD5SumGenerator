import os
import hashlib
import openpyxl

# Funktion, um die MD5-Summe einer Datei zu berechnen
def calculate_md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

# Funktion, um MD5-Summen für alle Dateien in einem Verzeichnis und Unterordnern zu erstellen
def calculate_md5_for_directory(directory):
    md5_data = {}
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            md5_data[file_path] = calculate_md5(file_path)
    return md5_data

# Funktion, um die MD5-Summen in eine Excel-Datei zu schreiben
def write_md5_to_excel(md5_data, excel_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Datei", "MD5-Summe"])
    for file_path, md5_sum in md5_data.items():
        ws.append([file_path, md5_sum])
    wb.save(excel_file)

if __name__ == "__main__":
    directory_path = "C:\Projekte\K016_Sicherheits-SW"  # Hier das Verzeichnis angeben, das du durchsuchen möchtest
    excel_file_path = "md5_summen.xlsx"  # Name der Excel-Datei, in der die Ergebnisse gespeichert werden sollen

    md5_data = calculate_md5_for_directory(directory_path)
    write_md5_to_excel(md5_data, excel_file_path)