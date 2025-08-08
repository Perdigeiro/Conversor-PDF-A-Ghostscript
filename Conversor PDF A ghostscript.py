from tkinter import filedialog, messagebox, Tk, StringVar, Label, Button
import subprocess
import os

def converter_para_pdfa(input_pdf, output_pdf, versao_pdfa="1b"):
    """
    Converte um PDF padrão para PDF/A usando Ghostscript com perfil ICC sRGB.

    Args:
        input_pdf (str): Caminho completo do PDF original.
        output_pdf (str): Caminho onde o PDF/A convertido será salvo.
        versao_pdfa (str): Versão do PDF/A ("1b", "2b" ou "3b").
    """

    gs_executable = r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"
    icc_profile = r"sRGB.icc"

    pdfa_dict = {
        "1b": "-dPDFA=1",
        "2b": "-dPDFA=2",
        "3b": "-dPDFA=3"
    }

    if versao_pdfa not in pdfa_dict:
        print("Versão PDF/A inválida. Use '1b', '2b' ou '3b'.")
        return False

    if not os.path.exists(gs_executable):
        print(f"[ERRO] Ghostscript não encontrado em: {gs_executable}")
        return False

    if not os.path.exists(icc_profile):
        print(f"[ERRO] Perfil ICC não encontrado em: {icc_profile}")
        return False

    if not os.path.exists(input_pdf):
        print(f"[ERRO] Arquivo de entrada não encontrado: {input_pdf}")
        return False

    comando = [
        gs_executable,
        pdfa_dict[versao_pdfa],
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-dUseCIEColor",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=UseDeviceIndependentColor",
        "-dPDFACompatibilityPolicy=1",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={output_pdf}",
        f"-sDefaultRGBProfile={icc_profile}",
        input_pdf
    ]

    print(f"[INFO] Iniciando conversão para PDF/A-{versao_pdfa.upper()}")
    print(f"Entrada: {input_pdf}")
    print(f"Saída:   {output_pdf}")

    try:
        resultado = subprocess.run(comando, check=True, capture_output=True, text=True)
        print("[OK] Conversão concluída com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print("[ERRO] Ghostscript retornou erro:")
        print(e.stderr)
        return False
    except Exception as e:
        print(f"[ERRO] Falha inesperada: {e}")
        return False


if __name__ == "__main__":
    input_pdf = r"C:\Users\matheus.augusto\Downloads\teste.pdf"
    output_pdf = r"C:\Users\matheus.augusto\Downloads\Log_PDFA.pdf"
    converter_para_pdfa(input_pdf, output_pdf, versao_pdfa="1b")