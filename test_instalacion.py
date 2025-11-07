#!/usr/bin/env python3
"""Test de instalación del sistema"""

def test_imports():
    """Verifica que todas las librerías estén instaladas"""
    print("Verificando importaciones...")
    
    try:
        import PyPDF2
        print("✓ PyPDF2")
    except ImportError:
        print("✗ PyPDF2 - ERROR")
        return False
    
    try:
        from PIL import Image
        print("✓ Pillow")
    except ImportError:
        print("✗ Pillow - ERROR")
        return False
    
    try:
        import numpy
        print("✓ NumPy")
    except ImportError:
        print("✗ NumPy - ERROR")
        return False
    
    try:
        from reportlab.lib.pagesizes import A4
        print("✓ ReportLab")
    except ImportError:
        print("✗ ReportLab - ERROR")
        return False
    
    try:
        import pytesseract
        print("✓ pytesseract")
    except ImportError:
        print("✗ pytesseract - ERROR")
        return False
    
    try:
        import language_tool_python
        print("✓ language-tool-python")
    except ImportError:
        print("✗ language-tool-python - ERROR")
        return False
    
    return True

if __name__ == "__main__":
    print("=" * 50)
    print("TEST DE INSTALACIÓN")
    print("=" * 50 + "\n")
    
    if test_imports():
        print("\n✓ Todas las dependencias están instaladas correctamente")
        print("✓ El sistema está listo para usar")
    else:
        print("\n✗ Faltan dependencias")
        print("Ejecuta: pip install -r requirements.txt")
