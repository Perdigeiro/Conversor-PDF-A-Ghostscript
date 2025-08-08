import os
import subprocess
import tempfile
import win32com.client
from tkinter import filedialog, Tk

def imprimir_pdf_como_pdfa_temporario(input_path):
    """
    Reimprime o PDF via Word como PDF/A (ISO19005-1) e salva em arquivo temporário.
    Retorna o caminho do PDF gerado.
    """
    input_path = os.path.abspath(input_path)
    input_path = input_path.replace("/", "\\")  # Corrige caminhos para o formato do Word

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    temp_pdfa = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_path = temp_pdfa.name
    temp_pdfa.close()

    doc = None
    try:
        doc = word.Documents.Open(input_path, ReadOnly=1)
        doc.ExportAsFixedFormat(
            OutputFileName=temp_path,
            ExportFormat=17,  # PDF
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=True  # PDF/A
        )
        print(f"[OK] PDF impresso como PDF/A: {temp_path}")
        return temp_path
    except Exception as e:
        print(f"[ERRO] Falha ao imprimir via Word: {e}")
        return None
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def converter_para_pdfa(input_pdf, output_pdf, versao_pdfa="1b"):
    """
    Converte um PDF padrão para PDF/A usando Ghostscript com perfil ICC sRGB.
    """
    gs_executable = r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"
    icc_profile = r"sRGB.icc"

    pdfa_dict = {
        "1b": "-dPDFA=1",
        "2b": "-dPDFA=2",
        "3b": "-dPDFA=3"
    }

    comando = [
        gs_executable,
        pdfa_dict.get(versao_pdfa, "-dPDFA=1"),
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-dUseCIEColor",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=UseDeviceIndependentColor",
        "-dPDFACompatibilityPolicy=1",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=false",
        "-dPDFSETTINGS=/prepress",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={output_pdf}",
        f"-sDefaultRGBProfile={icc_profile}",
        input_pdf
    ]

    try:
        resultado = subprocess.run(comando, check=True, capture_output=True, text=True)
        print("[OK] Conversão para PDF/A concluída com sucesso.")
        return True
    except subprocess.CalledProcessError as e:
        print("[ERRO] Ghostscript falhou:")
        print(e.stderr)
        return False

def selecionar_e_converter():
    Tk().withdraw()
    input_pdf = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")],
        title="Selecione o PDF original"
    )
    if not input_pdf:
        print("[INFO] Nenhum arquivo selecionado.")
        return

    output_pdf = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF/A files", "*.pdf")],
        title="Salvar como PDF/A"
    )
    if not output_pdf:
        print("[INFO] Salvamento cancelado.")
        return

    # 1. Imprimir para PDF/A temporário
    pdf_temp = imprimir_pdf_como_pdfa_temporario(input_pdf)
    if not pdf_temp or not os.path.exists(pdf_temp):
        print("[ERRO] Falha na impressão temporária.")
        return

    try:
        # 2. Converter o PDF temporário em PDF/A validado com Ghostscript
        sucesso = converter_para_pdfa(pdf_temp, output_pdf)
        if sucesso:
            print(f"[SUCESSO] Arquivo final salvo em: {output_pdf}")
        else:
            print("[ERRO] Falha na conversão final.")
    finally:
        # 3. Apagar o temporário
        if os.path.exists(pdf_temp):
            os.remove(pdf_temp)
            print(f"[INFO] Arquivo temporário removido: {pdf_temp}")

if __name__ == "__main__":
    selecionar_e_converter()
