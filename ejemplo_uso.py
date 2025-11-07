#!/usr/bin/env python3
"""
Ejemplo de uso del Sistema RPA de Verificación de Admisibilidad
"""

import sys
import os
from pathlib import Path

# Agregar el directorio src al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from verificador_admisibilidad import ejecutar_verificacion_rpa

def main():
    print("=" * 80)
    print("SISTEMA RPA DE VERIFICACIÓN DE ADMISIBILIDAD DOCUMENTAL")
    print("=" * 80)
    print()
    
    # Buscar PDFs en la carpeta de entrada
    carpeta_entrada = Path("input/documentos_pendientes")
    archivos_pdf = list(carpeta_entrada.glob("*.pdf"))
    
    # Verificar si hay PDFs
    if not archivos_pdf:
        print("⚠️  No se encontraron archivos PDF en: input/documentos_pendientes/")
        print()
        print("Por favor:")
        print("1. Coloca tus documentos PDF en la carpeta: input/documentos_pendientes/")
        print("2. Ejecuta nuevamente este script")
        print()
        return
    
    # Mostrar archivos encontrados
    print(f"📄 Se encontraron {len(archivos_pdf)} documento(s) PDF:\n")
    for i, archivo in enumerate(archivos_pdf, 1):
        print(f"   {i}. {archivo.name}")
    print()
    
    # Si hay un solo archivo, procesarlo directamente
    if len(archivos_pdf) == 1:
        documento = archivos_pdf[0]
        print(f"→ Procesando automáticamente: {documento.name}")
        print()
    else:
        # Permitir al usuario elegir
        print("Opciones:")
        print("  - Ingresa un número (1-{}) para procesar un documento específico".format(len(archivos_pdf)))
        print("  - Ingresa 'todos' para procesar todos los documentos")
        print("  - Ingresa 'salir' para cancelar")
        print()
        
        opcion = input("Tu elección: ").strip().lower()
        
        if opcion == 'salir':
            print("Operación cancelada.")
            return
        elif opcion == 'todos':
            # Procesar todos los documentos
            print()
            print("=" * 80)
            print("PROCESANDO TODOS LOS DOCUMENTOS")
            print("=" * 80)
            print()
            
            for i, archivo in enumerate(archivos_pdf, 1):
                print(f"\n[{i}/{len(archivos_pdf)}] Procesando: {archivo.name}")
                print("-" * 80)
                
                reporte = f"output/reportes_pdf/reporte_{archivo.stem}.pdf"
                
                try:
                    resultados = ejecutar_verificacion_rpa(str(archivo), reporte)
                    if resultados:
                        print(f"✓ Reporte generado: {reporte}")
                    else:
                        print(f"✗ Error al procesar: {archivo.name}")
                except Exception as e:
                    print(f"✗ Error: {e}")
            
            print()
            print("=" * 80)
            print("PROCESO COMPLETADO")
            print("=" * 80)
            print(f"✓ Revisa los reportes en: output/reportes_pdf/")
            return
        else:
            # Procesar un documento específico
            try:
                indice = int(opcion) - 1
                if 0 <= indice < len(archivos_pdf):
                    documento = archivos_pdf[indice]
                    print(f"\n→ Procesando: {documento.name}")
                    print()
                else:
                    print("✗ Número inválido. Operación cancelada.")
                    return
            except ValueError:
                print("✗ Opción inválida. Operación cancelada.")
                return
    
    # Ruta de salida del reporte
    reporte = f"output/reportes_pdf/reporte_{documento.stem}.pdf"
    
    # Ejecutar verificación
    print("Iniciando verificación...")
    print("-" * 80)
    
    try:
        resultados = ejecutar_verificacion_rpa(str(documento), reporte)
        
        if resultados:
            print()
            print("=" * 80)
            print("✓ VERIFICACIÓN COMPLETADA EXITOSAMENTE")
            print("=" * 80)
            print(f"✓ Reporte PDF: {reporte}")
            print(f"✓ Reporte JSON: {reporte.replace('.pdf', '.json')}")
            print()
        else:
            print()
            print("✗ Ocurrió un error durante la verificación")
    except Exception as e:
        print()
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
