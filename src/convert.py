# Importing the Pandas library to read and manipulate the Excel file;
import pandas as pd

# Class responsible for converting Excel files to Markdown;
class Convert:
    def __init__(self:object, input_file_path: str, output_file_path: str) -> None:
        """
        Inicializa a classe Convert.
        
        Args:
            input_file_path (str): Caminho do arquivo Excel (.xlsx).
            output_file_path (str): Caminho de saída para o arquivo Markdown (.md).

        Returns:
            None
        """
        self.input_file = input_file_path
        self.output_file = output_file_path

        # Load Excel file;
        self.xlsx = pd.ExcelFile(self.input_file)

        # Read all spreadsheets from Excel file;
        self.sheet_names = self.xlsx.sheet_names
    def excel_to_markdown(self: object) -> str:
        """
        Converte todas as planilhas de uma pasta Excel para Markdown e salva em um arquivo separando por '---'.

        Returns:
            str: Conteúdo da tabela em formato Markdown.
        """
        # The markdown_content is the variable to be used to store the final output;
        markdown_content = ""
        # Loop through all sheet names and convert each to Markdown;
        for spreadsheet in self.sheet_names:
            # Read the current spreadsheet into a Pandas DataFrame;
            data_frame = pd.read_excel(self.input_file, sheet_name=spreadsheet)
            # Add a title for the current spreadsheet;
            markdown_content += f"## {spreadsheet.capitalize()}\n\n"
            # Convert the DataFrame to Markdown format;
            markdown_content += data_frame.to_markdown(index=False) + "\n\n"
            # Add a separator between sheets;
            markdown_content += "---\n\n"

        # Save the final Markdown content to a file;
        with open(self.output_file, 'w', encoding='utf-8') as md_file:
            md_file.write(markdown_content)
            
        # Return the final Markdown content;
        return markdown_content