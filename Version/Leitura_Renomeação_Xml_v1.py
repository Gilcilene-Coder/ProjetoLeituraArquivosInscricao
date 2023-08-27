import os
import zipfile
import re
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, Toplevel, OptionMenu


def main(zip_path, output_dir, mode):
    # Contador de erros
    error_count = 0

    # Índice para selecionar o IE (0 para entrada, 1 para saída)
    ie_index = 0 if mode.lower() == "entrada" else 1

    # Abrir o arquivo zip
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # Iterar sobre cada arquivo dentro do zip
        files = zip_ref.namelist()
        files_content = defaultdict(list)
        for file in files:
            # Processa apenas arquivos .xml
            if file.endswith(".xml"):
                try:
                    # Abrir cada arquivo
                    with zip_ref.open(file) as f:
                        content = f.read().decode('utf-8')  # Ler o conteúdo do arquivo como texto

                        # Encontrar a palavra <IE> e pegar os números após o >
                        matches = re.findall(r'<IE>(\d+)', content)
                        if len(matches) <= ie_index:
                            continue
                        ie_number = matches[ie_index]

                        # Extrair chave de acesso do nome do arquivo
                        chave_acesso = file.split('-')[-1].split('.')[0]

                        # Armazenar conteúdo
                        files_content[ie_number].append((chave_acesso, content))

                except Exception as e:
                    print(f"Não foi possível ler o arquivo: {file}. Erro: {e}")
                    error_count += 1
                    continue

        # Criar diretórios e salvar arquivos
        for ie_number, contents in files_content.items():
            folder_path = os.path.join(output_dir, ie_number)
            os.makedirs(folder_path, exist_ok=True)
            for chave_acesso, content in contents:
                with open(os.path.join(folder_path, f"{chave_acesso}.xml"), "w") as f:
                    f.write(content)

    # Exibir mensagem de status
    if error_count == 0:
        messagebox.showinfo("Status", "Processo concluído com sucesso!")
    else:
        messagebox.showwarning("Status",
                               f"Processo concluído com {error_count} erro(s). Verifique o console para detalhes.")


if __name__ == "__main__":
    # Criar uma janela de seleção de arquivo
    root = tk.Tk()
    root.withdraw()

    # Criar uma janela para seleção de tipo de nota
    select_window = Toplevel(root)
    select_window.title("Leitura e Renomeação de Xml")
    select_window.geometry("350x150+400+400")

    label = tk.Label(select_window, text="Selecione o tipo de nota que deseja fazer a renomeção:")
    label.pack(pady=10)

    var = StringVar(select_window)
    var.set("")  # Deixar em branco inicialmente

    options = OptionMenu(select_window, var, "Entrada", "Saída")
    options.pack(pady=10)

    def on_submit():
        mode = var.get().lower()
        if mode == "":
            messagebox.showerror("Erro", "Selecione o tipo de nota.")
            return

        select_window.destroy()

        zip_path = filedialog.askopenfilename(title="Selecione o arquivo ZIP", filetypes=[("Arquivos ZIP", "*.zip")])
        output_dir = filedialog.askdirectory(title="Selecione o diretório de saída")

        if zip_path and output_dir:
            main(zip_path, output_dir, mode)
        else:
            print("Seleção de arquivo ou diretório cancelada.")

    submit_button = tk.Button(select_window, text="Executar", command=on_submit)
    submit_button.pack(pady=10)

    select_window.mainloop()
