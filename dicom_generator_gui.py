import os
import pydicom
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pydicom.uid import generate_uid

def select_dicom():
    file_path = filedialog.askopenfilename(filetypes=[("DICOM Files", "*.dcm")])
    dicom_entry.delete(0, tk.END)
    dicom_entry.insert(0, file_path)

def select_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, folder_path)

def generate_dicoms():
    dicom_path = dicom_entry.get()
    excel_path = excel_entry.get()
    output_folder = output_entry.get()

    if not os.path.exists(dicom_path):
        messagebox.showerror("Erro", "Arquivo DICOM base não encontrado!")
        return
    if not os.path.exists(excel_path):
        messagebox.showerror("Erro", "Arquivo Excel não encontrado!")
        return
    if not os.path.exists(output_folder):
        messagebox.showerror("Erro", "Pasta de saída inválida!")
        return

    df = pd.read_excel(excel_path)
    log = []

    for index, row in df.iterrows():
        try:
            ds = pydicom.dcmread(dicom_path)
            ds.PatientName = row["PatientName"]
            ds.PatientID = str(row["PatientID"])
            ds.StudyDate = str(row["StudyDate"])

            ds.StudyInstanceUID = generate_uid() if row["StudyInstanceUID"] == "AUTO" else row["StudyInstanceUID"]
            ds.SeriesInstanceUID = generate_uid() if row["SeriesInstanceUID"] == "AUTO" else row["SeriesInstanceUID"]
            ds.SOPInstanceUID = generate_uid() if row["SOPInstanceUID"] == "AUTO" else row["SOPInstanceUID"]

            filename = f"{output_folder}/{ds.SOPInstanceUID}.dcm"
            ds.save_as(filename)
            log.append(f"[SUCESSO] {filename} gerado com sucesso!")
        except Exception as e:
            log.append(f"[ERRO] Falha ao processar {row.get('PatientID', 'Desconhecido')}: {str(e)}")

    with open(os.path.join(output_folder, "relatorio.txt"), "w", encoding="utf-8") as log_file:
        log_file.write("\n".join(log))

    messagebox.showinfo("Concluído", "Processo finalizado! Relatório salvo na pasta de saída.")

# Criar interface gráfica
root = tk.Tk()
root.title("Gerador de DICOMs")
root.geometry("500x300")

tk.Label(root, text="Arquivo DICOM base:").pack()
dicom_entry = tk.Entry(root, width=50)
dicom_entry.pack()
tk.Button(root, text="Selecionar", command=select_dicom).pack()

tk.Label(root, text="Arquivo Excel de dados:").pack()
excel_entry = tk.Entry(root, width=50)
excel_entry.pack()
tk.Button(root, text="Selecionar", command=select_excel).pack()

tk.Label(root, text="Pasta de saída:").pack()
output_entry = tk.Entry(root, width=50)
output_entry.pack()
tk.Button(root, text="Selecionar", command=select_output_folder).pack()

tk.Button(root, text="Gerar DICOMs", command=generate_dicoms).pack(pady=10)
root.mainloop()
