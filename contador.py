import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter.ttk import Treeview, Style
from collections import defaultdict

def browse_directory():
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory()
    return directory

def count_files_and_folders(directory):
    folder_count = 0
    file_count = 0
    extension_count = defaultdict(int)
    folders_info = defaultdict(lambda: {"subdirs": [], "files": []})

    for root, dirs, files in os.walk(directory):
        folder_count += 1
        folders_info[root]["subdirs"] = dirs
        folders_info[root]["files"] = files
        for file in files:
            file_count += 1
            ext = os.path.splitext(file)[1].lower()
            extension_count[ext] += 1

    return folder_count, file_count, extension_count, folders_info

def format_results(tree, folders_info):
    def format_folder(parent, path):
        folder_id = tree.insert(parent, 'end', text=os.path.basename(path), values=("Folder", ""))
        for subdir in folders_info[path]["subdirs"]:
            subdir_path = os.path.join(path, subdir)
            format_folder(folder_id, subdir_path)
        for file in folders_info[path]["files"]:
            tree.insert(folder_id, 'end', text=file, values=("File", ""))

    root_dir = list(folders_info.keys())[0]
    format_folder("", root_dir)

def group_extensions(extension_count):
    grouped_extensions = defaultdict(list)
    for ext, count in extension_count.items():
        if ext in ['.xls', '.xlsx', '.xltx']:
            grouped_extensions['Excel'].append((ext, count))
        elif ext in ['.doc', '.docx']:
            grouped_extensions['Word'].append((ext, count))
        elif ext in ['.ppt', '.pptx']:
            grouped_extensions['PowerPoint'].append((ext, count))
        elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp']:
            grouped_extensions['Imagens'].append((ext, count))
        elif ext in ['.zip', '.rar', '.7z']:
            grouped_extensions['Arquivos Comprimidos'].append((ext, count))
        else:
            grouped_extensions['Outros'].append((ext, count))
    return grouped_extensions

def format_extension_results(grouped_extensions):
    from tabulate import tabulate
    result_lines = []
    col_count = 5

    for group, extensions in grouped_extensions.items():
        table = []
        current_line = []
        result_lines.append(f"{group}:\n")
        for ext, count in extensions:
            current_line.append(f"{ext}: {count} arquivo(s)")
            if len(current_line) == col_count:
                table.append(current_line)
                current_line = []
        if current_line:
            table.append(current_line)
        result_lines.append(tabulate(table, tablefmt="grid"))
        result_lines.append("")  # Add an empty line after each group

    return "\n\n".join(result_lines)

def main():
    directory = browse_directory()
    if not directory:
        print("Nenhum diretório selecionado.")
        return

    # Configuração da janela Tkinter
    root = tk.Tk()
    root.title("Resultados da Análise de Diretório")
    root.state('zoomed')  # Abre a janela maximizada

    style = Style(root)
    style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
    style.configure("Treeview", font=("Arial", 10))

    tree = Treeview(root, columns=("Type", "Details"), show='tree headings')
    tree.heading("#0", text="Nome")
    tree.heading("#1", text="Tipo")
    tree.heading("#2", text="Detalhes")
    tree.pack(expand=True, fill=tk.BOTH, pady=10)

    folder_count, file_count, extension_count, folders_info = count_files_and_folders(directory)

    result_text = f"Número total de pastas: {folder_count}\n"
    result_text += f"Número total de arquivos: {file_count}\n\n"
    format_results(tree, folders_info)

    grouped_extensions = group_extensions(extension_count)
    result_text += "Número de arquivos por extensão:\n\n"
    result_text += format_extension_results(grouped_extensions)

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=10)
    text_area.pack(pady=10, expand=True, fill=tk.BOTH)
    text_area.insert(tk.END, result_text)

    close_button = tk.Button(root, text="Fechar", command=root.quit)
    close_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
