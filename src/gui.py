# Importing necessary modules;
import os, sys

# Importing PyQt6 to create a graphical user interface;
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QMessageBox, QTextEdit
)

# Importing my modules;
from src.convert import Convert

# Cretate a class to manage the user interface;
class ConvertApp(QWidget):
    def __init__(self:object) -> None:
        # Initializing the QWidget class so that the current class inherits its properties;
        super().__init__()
        # Define the window Title;
        self.setWindowTitle("Conversor Excel → Markdown")
        # Define the window Dimensions;
        self.setGeometry(200, 200, 600, 400)
        self.layout = QVBoxLayout()

        # Labels;
        self.input_label = QLabel("Arquivo Excel: (não selecionado)")
        self.output_label = QLabel("Destino Markdown: (não selecionado)")

        # Buttons;
        self.btn_select_input = QPushButton("Selecionar Excel")
        self.btn_select_output = QPushButton("Selecionar Pasta de Saída")
        self.btn_convert = QPushButton("Converter")

        # Preview result;
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)

        # Adding Widgets to the Layout;
        self.layout.addWidget(self.input_label)
        self.layout.addWidget(self.btn_select_input)
        self.layout.addWidget(self.output_label)
        self.layout.addWidget(self.btn_select_output)
        self.layout.addWidget(self.btn_convert)
        self.layout.addWidget(QLabel("Prévia da Tabela em Markdown:"))
        self.layout.addWidget(self.preview)

        self.setLayout(self.layout)

        # Connecting Buttons to their Functions;
        self.btn_select_input.clicked.connect(self.select_input_file)
        self.btn_select_output.clicked.connect(self.select_output_folder)
        self.btn_convert.clicked.connect(self.convert_file)

        # Paths;
        self.input_file = None
        self.output_folder = None

    # Function to select the input file
    def select_input_file(self:object) -> None:
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Selecionar Arquivo Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            self.input_file = file_path
            self.input_label.setText(f"Arquivo Excel: {file_path}")

    # Function to select the output folder
    def select_output_folder(self:object) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Saída", "")
        if folder:
            self.output_folder = folder
            self.output_label.setText(f"Destino Markdown: {folder}")
    
    # Function to convert the file;
    def convert_file(self:object) -> None:
        if not self.input_file or not self.output_folder:
            QMessageBox.warning(self, "Erro", "Selecione um arquivo de entrada e uma pasta de saída!")
            return
        # Define the name to output file;
        base_name = os.path.splitext(os.path.basename(self.input_file))[0]
        output_file = os.path.join(self.output_folder, f"{base_name}.md")

        try:
            convert = Convert(self.input_file, output_file)
            markdown_content = convert.excel_to_markdown()
            self.preview.setText(markdown_content)

            QMessageBox.information(self, "Sucesso", f"Arquivo convertido com sucesso!\nSalvo em: {output_file}")
        except Exception as error:
            QMessageBox.critical(self, "Erro", f"Erro ao converter arquivo:\n{error}")

# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConvertApp()
    window.show()
    sys.exit(app.exec())