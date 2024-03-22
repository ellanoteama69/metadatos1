from PyPDF2 import PdfReader
from docx import Document
import openpyxl

def extract_pdf_metadata(file_path):
    with open(file_path, 'rb') as file:
        reader = PdfReader(file)
        metadata = reader.metadata
    return metadata

def extract_docx_metadata(file_path):
    doc = Document(file_path)
    metadata = {}
    core_properties = doc.core_properties
    metadata['Title'] = core_properties.title
    metadata['Author'] = core_properties.author
    metadata['Subject'] = core_properties.subject
    metadata['Keywords'] = core_properties.keywords
    metadata['Created'] = core_properties.created
    metadata['Modified'] = core_properties.modified
    return metadata

def extract_xlsx_metadata(file_path):
    wb = openpyxl.load_workbook(file_path)
    properties = wb.properties
    metadata = {
        'Title': properties.title,
        'Author': properties.creator,
        'Created': properties.created,
        'Modified': properties.modified
    }
    return metadata

def extract_metadata(file_type, file_path):
    switch = {
        'pdf': extract_pdf_metadata,
        'docx': extract_docx_metadata,
        'xlsx': extract_xlsx_metadata
    }
    return switch[file_type.lower()](file_path)

if __name__ == "__main__":
    file_type = input("Ingrese el tipo de archivo que desea extraer (pdf, docx, xlsx): ").lower()
    file_path = input("Ingrese la ruta del archivo: ")

    try:
        metadata = extract_metadata(file_type, file_path)
        print("Metadata:")
        print(metadata)
    except KeyError:
        print("Tipo de archivo no compatible. Por favor, seleccione 'pdf', 'docx' o 'xlsx'.")
    except FileNotFoundError:
        print("Archivo no encontrado. Verifique la ruta proporcionada.")
    except Exception as e:
        print("Ocurrió un error:", e)
