# Importing necessary modules;
import os, sys

# Importing PQt6 to create a Application;
from PyQt6.QtWidgets import QApplication

# Importing my modules;
from src.convert import Convert
from src.gui import ConvertApp
from src.utils import print_banner

def main() -> None:
    print_banner()
    while True:
        preferences = input("Deseja abrir a interface gráfica?['s' para abrir interface GUI, 'n' para continuar na interface CLI ou somente 'ENTER' para cancelar e sair]: ").strip().lower()
        if preferences == "s":
            app = QApplication(sys.argv)
            window = ConvertApp()
            window.show()
            sys.exit(app.exec())
            return
        elif preferences == "n":
            print("Abrindo interface CLI...")
            while True:
                input_file_path = input("Cole aqui o caminho do arquivo Excel que você deseja converter: ")
                output_folder_path = input("Cole aqui o caminho onde você deseja salvar o arquivo Markdown: ")
                if not input_file_path or not output_folder_path:
                    print("Caminho inválido. Tente novamente.")
                    continue
                try:
                    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
                    output_file_path = os.path.join(output_folder_path, f"{base_name}.md")
                    convert = Convert(input_file_path, output_file_path)
                    markdown_content = convert.excel_to_markdown()
                    print("Arquivo convertido com sucesso!")
                    print(f"Conteúdo Markdown:\n{markdown_content}")
                    break
                except Exception as error:
                    print(f"Erro ao converter arquivo: {error} \n\n Tente novamente.")
                    continue
            return
        elif preferences == "":
            sys.exit()
        else:
            print("Opção inválida. Tente novamente.")
            continue

# Run the application;
if __name__ == "__main__":
    main()