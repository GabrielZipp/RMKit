# ============================================
# RMKIT - SIMPLES E DIRETO (teste_pdf.py)
# ============================================

from docx import Document
from datetime import datetime
import os
import comtypes.client
import tempfile

# ===== SEUS DADOS (EDITAR AQUI) =====
dados = {
    '{{NOME}}': 'MARIA APARECIDA SANTOS',
    '{{NACIONALIDADE}}': 'brasileira',
    '{{ESTADO_CIVIL}}': 'solteira',
    '{{PROFISSAO}}': 'advogada',
    '{{RG}}': '98.765.432-1 SSP/SP',
    '{{CPF}}': '987.654.321-00',
    '{{RUA}}': 'Paulista',
    '{{NUMERO}}': '456',
    '{{BAIRRO}}': 'Jardins',
    '{{CEP}}': '01310-100',
    '{{CIDADE}}': 'São Paulo',
    '{{ESTADO}}': 'SP',
    '{{TIPO_ACAO}}': 'Cível'
}

# ===== CONFIGURAÇÕES =====
TEMPLATE = "procuração.docx"  # Seu arquivo modelo
PDF_SAIDA = "minha_procuracao.pdf"  # Nome do PDF gerado

# ============================================
# CÓDIGO - NÃO PRECISA MEXER DAQUI PRA BAIXO
# ============================================

print("🚀 INICIANDO...")

# Verificar se template existe
if not os.path.exists(TEMPLATE):
    print(f"❌ ERRO: Arquivo '{TEMPLATE}' não encontrado!")
    exit()

try:
    # 1. CRIAR PASTA TEMPORÁRIA
    temp_dir = tempfile.mkdtemp()
    temp_docx = os.path.join(temp_dir, "temp.docx")
    print("✅ Pasta temporária criada")
    
    # 2. ABRIR TEMPLATE
    doc = Document(TEMPLATE)
    print("✅ Template aberto")
    
    # 3. ADICIONAR DATA AUTOMÁTICA
    meses = {1:'janeiro',2:'fevereiro',3:'março',4:'abril',5:'maio',6:'junho',
             7:'julho',8:'agosto',9:'setembro',10:'outubro',11:'novembro',12:'dezembro'}
    hoje = datetime.now()
    dados['{{DATA_EXTENSO}}'] = f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"
    print("✅ Data gerada")
    
    # 4. SUBSTITUIR VARIÁVEIS
    for paragraph in doc.paragraphs:
        for var, valor in dados.items():
            if var in paragraph.text:
                for run in paragraph.runs:
                    if var in run.text:
                        run.text = run.text.replace(var, str(valor))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for var, valor in dados.items():
                        if var in paragraph.text:
                            for run in paragraph.runs:
                                if var in run.text:
                                    run.text = run.text.replace(var, str(valor))
    print("✅ Variáveis substituídas")
    
    # 5. SALVAR DOCX TEMPORÁRIO
    doc.save(temp_docx)
    print("✅ DOCX temporário salvo")
    
    # 6. CONVERTER PARA PDF (usa Word)
    print("🔄 Convertendo para PDF (aguarde)...")
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    doc_word = word.Documents.Open(temp_docx)
    doc_word.SaveAs(PDF_SAIDA, FileFormat=17)  # 17 = PDF
    doc_word.Close()
    word.Quit()
    print("✅ Conversão concluída")
    
    # 7. LIMPAR ARQUIVOS TEMPORÁRIOS
    os.remove(temp_docx)
    os.rmdir(temp_dir)
    
    print(f"\n✨ SUCESSO! PDF gerado: {PDF_SAIDA}")
    print(f"📁 Caminho: {os.path.abspath(PDF_SAIDA)}")
    
except Exception as e:
    print(f"\n❌ ERRO: {e}")