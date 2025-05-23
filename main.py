import os
import sys
import locale
from docx2pdf import convert

# Dictionary of messages translations
MESSAGES = {
    "path_not_exist": {
        "en": "Error: The path '{path}' does not exist.",
        "es": "Error: La ruta '{path}' no existe.",
        "fr": "Erreur : Le chemin '{path}' n'existe pas.",
        "it": "Errore: Il percorso '{path}' non esiste."
    },
    "converted_file": {
        "en": "Converted '{path}' to PDF successfully.",
        "es": "Convertido '{path}' a PDF con éxito.",
        "fr": "Converti '{path}' en PDF avec succès.",
        "it": "convertito con successo '{Path}' in pdf"
    },
    "failed_convert_file": {
        "en": "Failed to convert '{path}': {error}",
        "es": "Error al convertir '{path}': {error}",
        "fr": "Échec de la conversion de '{path}' : {error}",
        "it": "Errore durante la conversione di '{path}': {error}"
    },
    "file_not_docx": {
        "en": "Error: The file is not a .docx file.",
        "es": "Error: El archivo no es un archivo .docx.",
        "fr": "Erreur : Le fichier n'est pas un fichier .docx.",
        "it": "Errore: Il file non è un file .docx."
    },
    "converted_dir": {
        "en": "Converted all .docx files in directory '{path}' to PDF successfully.",
        "es": "Convertidos todos los archivos .docx en el directorio '{path}' a PDF con éxito.",
        "fr": "Tous les fichiers .docx dans le répertoire '{path}' ont été convertis en PDF avec succès.",
        "it": "Convertiti con successo tutti i file .docx nel percorso '{path}' in pdf"
    },
    "failed_convert_dir": {
        "en": "Failed to convert files in directory '{path}': {error}",
        "es": "Error al convertir archivos en el directorio '{path}': {error}",
        "fr": "Échec de la conversion des fichiers dans le répertoire '{path}' : {error}",
        "it": "Errore durante la conversione dei file nel percorso '{path}': {error}"
    },
    "input_not_file_or_dir": {
        "en": "Error: The input path is neither a file nor a directory.",
        "es": "Error: La ruta de entrada no es un archivo ni un directorio.",
        "fr": "Erreur : Le chemin d'entrée n'est ni un fichier ni un répertoire.",
        "it": "Errore: La cartella di input non è né un file né un percorso"
    },
    "usage": {
        "en": "Usage: python word_to_pdf.py <path_to_docx_or_directory> [--saveat <output_directory>]",
        "es": "Uso: python word_to_pdf.py <ruta_a_docx_o_directorio> [--saveat <directorio_de_salida>]",
        "fr": "Utilisation : python word_to_pdf.py <chemin_vers_docx_ou_répertoire> [--saveat <répertoire_de_sortie>]",
        "it": "Uso: python word_to_pdf.py <percorso_a_docx_o_cartella> [--saveat <percorso_di_output>] "
    }
}

def get_system_language():
    lang, _ = locale.getlocale()
    if lang:
        return lang.split('_')[0]
    return "en"

def translate(message_key, **kwargs):
    lang = get_system_language()
    if lang not in MESSAGES[message_key]:
        lang = "en"
    return MESSAGES[message_key][lang].format(**kwargs)

def print_translated(message_key, **kwargs):
    print(translate(message_key, **kwargs))

def convert_word_to_pdf(input_path, output_dir=None):
    """
    Convert Word documents to PDF.
    If input_path is a file, convert that file.
    If input_path is a directory, convert all .docx files in it.
    If output_dir is specified, save PDFs to that directory.
    """ 
    if not os.path.exists(input_path):
        print_translated("path_not_exist", path=input_path)
        return

    if output_dir and not os.path.exists(output_dir):
        print_translated("path_not_exist", path=output_dir)
        return

    if os.path.isfile(input_path):
        if input_path.lower().endswith(".docx"):
            try:
                if output_dir:
                    convert(input_path, output_dir)
                else:
                    convert(input_path)
                print_translated("converted_file", path=input_path)
            except Exception as e:
                print_translated("failed_convert_file", path=input_path, error=e)
        else:
            print_translated("file_not_docx")
    elif os.path.isdir(input_path):
        try:
            if output_dir:
                convert(input_path, output_dir)
            else:
                convert(input_path)
            print_translated("converted_dir", path=input_path)
        except Exception as e:
            print_translated("failed_convert_dir", path=input_path, error=e)
    else:
        print_translated("input_not_file_or_dir")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print_translated("usage")
    else:
        input_path = None
        output_dir = None
        args = sys.argv[1:]
        if "--saveat" in args:
            saveat_index = args.index("--saveat")
            if saveat_index == 0:
                print_translated("usage")
                sys.exit(1)
            input_path = args[0]
            if len(args) > saveat_index + 1:
                output_dir = args[saveat_index + 1]
            else:
                print_translated("usage")
                sys.exit(1)
        else:
            input_path = args[0]
        convert_word_to_pdf(input_path, output_dir)
