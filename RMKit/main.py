# ============================================
# RMKIT - ARQUIVO PRINCIPAL (main.py)
# ============================================

import tkinter as tk
import os
import sys

# Garantir que estamos no diretório certo
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Importar a interface
from interface import InterfaceRMKit

if __name__ == "__main__":
    # Verificar dependências
    try:
        import docx
        import comtypes
        import PyPDF2
    except ImportError as e:
        print(f"❌ Erro: Biblioteca não instalada - {e}")
        print("\nExecute: pip install python-docx comtypes PyPDF2")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    
    # Iniciar aplicação
    root = tk.Tk()
    app = InterfaceRMKit(root)
    root.mainloop()