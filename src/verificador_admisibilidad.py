"""
Sistema RPA de Verificación de Admisibilidad Documental
Autor: Sistema Automatizado
Fecha: 2025
"""

import os
import re
import hashlib
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple
from dataclasses import dataclass, field
import json

try:
    import PyPDF2
    from PIL import Image
    import numpy as np
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.units import inch
    import pytesseract
    from difflib import SequenceMatcher
    import language_tool_python
except ImportError as e:
    print(f"Error: Instalar dependencias faltantes - {e}")
    print("Ejecuta: pip install -r requirements.txt")
    exit(1)


@dataclass
class ResultadoVerificacion:
    """Estructura para almacenar resultados de verificación"""
    tipo_verificacion: str
    estado: str  # APROBADO, OBSERVADO, RECHAZADO
    detalles: List[str] = field(default_factory=list)
    porcentaje_cumplimiento: float = 0.0
    folios_afectados: List[int] = field(default_factory=list)


class VerificadorAdmisibilidad:
    """Clase principal para verificación de admisibilidad documental"""
    
    def __init__(self, ruta_documento: str, config: Dict = None):
        self.ruta_documento = ruta_documento
        self.config = config or self._config_default()
        self.resultados = []
        self.documento_pdf = None
        self.total_folios = 0
        
    def _config_default(self) -> Dict:
        """Configuración por defecto del sistema"""
        return {
            'umbral_blanco': 0.98,
            'umbral_ilegibilidad': 0.60,
            'umbral_plagio': 0.85,
            'verificar_ortografia': True,
            'idioma_ortografia': 'es',
            'formato_fecha': '%Y-%m-%d %H:%M:%S'
        }
    
    def ejecutar_verificacion_completa(self) -> Dict:
        """Ejecuta todas las verificaciones del documento"""
        print(f"Iniciando verificación de: {self.ruta_documento}")
        print(f"Fecha y hora: {datetime.now().strftime(self.config['formato_fecha'])}")
        
        if not self._cargar_documento():
            return self._generar_resultado_error()
        
        self._verificar_hojas_blanco()
        self._verificar_foliacion()
        self._verificar_folios_duplicados()
        self._verificar_ilegibilidad()
        self._verificar_plagio()
        
        if self.config['verificar_ortografia']:
            self._verificar_ortografia()
        
        return self._preparar_resultados()
    
    def _cargar_documento(self) -> bool:
        """Carga el documento PDF para procesamiento"""
        try:
            self.documento_pdf = PyPDF2.PdfReader(self.ruta_documento)
            self.total_folios = len(self.documento_pdf.pages)
            print(f"Documento cargado correctamente. Total de folios: {self.total_folios}")
            return True
        except Exception as e:
            resultado = ResultadoVerificacion(
                tipo_verificacion="Carga de Documento",
                estado="RECHAZADO",
                detalles=[f"Error al cargar documento: {str(e)}"]
            )
            self.resultados.append(resultado)
            return False
    
    def _verificar_hojas_blanco(self):
        """1.1 Verificación de hojas en blanco"""
        print("\n[1.1] Verificando hojas en blanco...")
        hojas_blanco = []
        
        for num_pagina in range(self.total_folios):
            try:
                pagina = self.documento_pdf.pages[num_pagina]
                texto = pagina.extract_text().strip()
                
                if len(texto) < 10:
                    hojas_blanco.append(num_pagina + 1)
            except Exception as e:
                print(f"Error en folio {num_pagina + 1}: {e}")
        
        estado = "APROBADO" if len(hojas_blanco) == 0 else "OBSERVADO"
        detalles = []
        
        if hojas_blanco:
            detalles.append(f"Se detectaron {len(hojas_blanco)} hojas en blanco")
            detalles.append(f"Folios afectados: {', '.join(map(str, hojas_blanco))}")
        else:
            detalles.append("No se detectaron hojas en blanco")
        
        resultado = ResultadoVerificacion(
            tipo_verificacion="1.1 Hojas en Blanco",
            estado=estado,
            detalles=detalles,
            porcentaje_cumplimiento=((self.total_folios - len(hojas_blanco)) / self.total_folios * 100),
            folios_afectados=hojas_blanco
        )
        self.resultados.append(resultado)
    
    def _verificar_foliacion(self):
        """1.2 Verificación de correlativo de foliación"""
        print("\n[1.2] Verificando foliación correlativa...")
        folios_encontrados = []
        folios_incorrectos = []
        
        for num_pagina in range(self.total_folios):
            try:
                pagina = self.documento_pdf.pages[num_pagina]
                texto = pagina.extract_text()
                
                patrones = [
                    r'folio\s*:?\s*(\d+)',
                    r'foja\s*:?\s*(\d+)',
                    r'página\s*:?\s*(\d+)',
                    r'^(\d+)$'
                ]
                
                folio_encontrado = None
                for patron in patrones:
                    match = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
                    if match:
                        folio_encontrado = int(match.group(1))
                        break
                
                folios_encontrados.append({
                    'pagina': num_pagina + 1,
                    'folio': folio_encontrado
                })
                
                if folio_encontrado != num_pagina + 1:
                    folios_incorrectos.append({
                        'pagina': num_pagina + 1,
                        'esperado': num_pagina + 1,
                        'encontrado': folio_encontrado
                    })
            except Exception as e:
                print(f"Error procesando folio {num_pagina + 1}: {e}")
        
        if len(folios_incorrectos) == 0:
            estado = "APROBADO"
            detalles = ["La foliación es correlativa y correcta"]
        elif len(folios_incorrectos) <= self.total_folios * 0.1:
            estado = "OBSERVADO"
            detalles = [f"Se detectaron {len(folios_incorrectos)} inconsistencias en la foliación"]
        else:
            estado = "RECHAZADO"
            detalles = [f"Foliación incorrecta en {len(folios_incorrectos)} folios"]
        
        for inc in folios_incorrectos[:10]:
            detalles.append(
                f"Página {inc['pagina']}: esperado folio {inc['esperado']}, encontrado {inc['encontrado']}"
            )
        
        if len(folios_incorrectos) > 10:
            detalles.append(f"... y {len(folios_incorrectos) - 10} inconsistencias más")
        
        porcentaje = ((self.total_folios - len(folios_incorrectos)) / self.total_folios * 100)
        
        resultado = ResultadoVerificacion(
            tipo_verificacion="1.2 Foliación Correlativa",
            estado=estado,
            detalles=detalles,
            porcentaje_cumplimiento=porcentaje,
            folios_afectados=[inc['pagina'] for inc in folios_incorrectos]
        )
        self.resultados.append(resultado)
    
    def _verificar_folios_duplicados(self):
        """1.2 Verificación de folios duplicados mediante hash"""
        print("\n[1.2] Verificando folios duplicados...")
        hashes_folios = {}
        duplicados = []
        
        for num_pagina in range(self.total_folios):
            try:
                pagina = self.documento_pdf.pages[num_pagina]
                texto = pagina.extract_text()
                
                hash_contenido = hashlib.md5(texto.encode()).hexdigest()
                
                if hash_contenido in hashes_folios:
                    duplicados.append({
                        'folio_original': hashes_folios[hash_contenido],
                        'folio_duplicado': num_pagina + 1
                    })
                else:
                    hashes_folios[hash_contenido] = num_pagina + 1
            except Exception as e:
                print(f"Error procesando folio {num_pagina + 1}: {e}")
        
        if len(duplicados) == 0:
            estado = "APROBADO"
            detalles = ["No se detectaron folios duplicados"]
        else:
            estado = "RECHAZADO"
            detalles = [f"ALERTA: Se detectaron {len(duplicados)} folios duplicados"]
            for dup in duplicados[:10]:
                detalles.append(
                    f"Folio {dup['folio_duplicado']} es idéntico al folio {dup['folio_original']}"
                )
        
        resultado = ResultadoVerificacion(
            tipo_verificacion="1.2 Folios Duplicados",
            estado=estado,
            detalles=detalles,
            porcentaje_cumplimiento=((self.total_folios - len(duplicados)) / self.total_folios * 100),
            folios_afectados=[dup['folio_duplicado'] for dup in duplicados]
        )
        self.resultados.append(resultado)
    
    def _verificar_ilegibilidad(self):
        """1.3 Verificación de ilegibilidad de información"""
        print("\n[1.3] Verificando ilegibilidad...")
        folios_ilegibles = []
        
        for num_pagina in range(self.total_folios):
            try:
                pagina = self.documento_pdf.pages[num_pagina]
                texto = pagina.extract_text()
                
                if len(texto) > 0:
                    caracteres_validos = sum(c.isalnum() or c.isspace() for c in texto)
                    porcentaje_legible = caracteres_validos / len(texto)
                    
                    if porcentaje_legible < self.config['umbral_ilegibilidad']:
                        folios_ilegibles.append({
                            'folio': num_pagina + 1,
                            'porcentaje': round(porcentaje_legible * 100, 2)
                        })
            except Exception as e:
                folios_ilegibles.append({
                    'folio': num_pagina + 1,
                    'porcentaje': 0,
                    'error': str(e)
                })
        
        umbral_pct = self.config['umbral_ilegibilidad'] * 100
        
        if len(folios_ilegibles) == 0:
            estado = "APROBADO"
            detalles = [f"Todos los folios superan el umbral de legibilidad ({umbral_pct}%)"]
        elif len(folios_ilegibles) <= self.total_folios * 0.05:
            estado = "OBSERVADO"
            detalles = [f"Se detectaron {len(folios_ilegibles)} folios con baja legibilidad"]
        else:
            estado = "RECHAZADO"
            detalles = [f"Cantidad significativa de folios ilegibles: {len(folios_ilegibles)}"]
        
        for ileg in folios_ilegibles[:10]:
            detalles.append(
                f"Folio {ileg['folio']}: {ileg['porcentaje']}% legible (mínimo requerido: {umbral_pct}%)"
            )
        
        resultado = ResultadoVerificacion(
            tipo_verificacion="1.3 Ilegibilidad de Información",
            estado=estado,
            detalles=detalles,
            porcentaje_cumplimiento=((self.total_folios - len(folios_ilegibles)) / self.total_folios * 100),
            folios_afectados=[ileg['folio'] for ileg in folios_ilegibles]
        )
        self.resultados.append(resultado)
    
    def _verificar_plagio(self):
        """1.3 Verificación de plagio entre secciones del documento"""
        print("\n[1.3] Verificando plagio...")
        secciones_texto = []
        casos_plagio = []
        
        for num_pagina in range(self.total_folios):
            try:
                pagina = self.documento_pdf.pages[num_pagina]
                texto = pagina.extract_text().strip()
                if len(texto) > 100:
                    secciones_texto.append({
                        'folio': num_pagina + 1,
                        'texto': texto
                    })
            except Exception as e:
                print(f"Error extrayendo texto del folio {num_pagina + 1}: {e}")
        
        for i in range(len(secciones_texto)):
            for j in range(i + 1, len(secciones_texto)):
                similitud = self._calcular_similitud(
                    secciones_texto[i]['texto'],
                    secciones_texto[j]['texto']
                )
                
                if similitud >= self.config['umbral_plagio']:
                    casos_plagio.append({
                        'folio_1': secciones_texto[i]['folio'],
                        'folio_2': secciones_texto[j]['folio'],
                        'similitud': round(similitud * 100, 2)
                    })
        
        if len(casos_plagio) == 0:
            estado = "APROBADO"
            detalles = ["No se detectaron casos de plagio interno"]
        else:
            estado = "OBSERVADO"
            detalles = [f"Se detectaron {len(casos_plagio)} casos de alta similitud"]
            for caso in casos_plagio[:10]:
                detalles.append(
                    f"Folios {caso['folio_1']} y {caso['folio_2']}: {caso['similitud']}% de similitud"
                )
        
        resultado = ResultadoVerificacion(
            tipo_verificacion="1.3 Verificación de Plagio",
            estado=estado,
            detalles=detalles,
            porcentaje_cumplimiento=100 - (len(casos_plagio) / max(len(secciones_texto), 1) * 100),
            folios_afectados=list(set([c['folio_1'] for c in casos_plagio] + [c['folio_2'] for c in casos_plagio]))
        )
        self.resultados.append(resultado)
    
    def _calcular_similitud(self, texto1: str, texto2: str) -> float:
        """Calcula similitud entre dos textos usando SequenceMatcher"""
        return SequenceMatcher(None, texto1, texto2).ratio()
    
    def _verificar_ortografia(self):
        """1.3 Verificación ortográfica del documento"""
        print("\n[1.3] Verificando ortografía...")
        try:
            tool = language_tool_python.LanguageTool(self.config['idioma_ortografia'])
            errores_por_folio = []
            total_errores = 0
            
            for num_pagina in range(min(self.total_folios, 50)):
                try:
                    pagina = self.documento_pdf.pages[num_pagina]
                    texto = pagina.extract_text()
                    
                    if len(texto.strip()) > 0:
                        matches = tool.check(texto)
                        if len(matches) > 0:
                            errores_por_folio.append({
                                'folio': num_pagina + 1,
                                'cantidad_errores': len(matches),
                                'ejemplos': [m.message[:80] for m in matches[:3]]
                            })
                            total_errores += len(matches)
                except Exception as e:
                    print(f"Error verificando ortografía en folio {num_pagina + 1}: {e}")
            
            tool.close()
            
            if total_errores == 0:
                estado = "APROBADO"
                detalles = ["No se detectaron errores ortográficos significativos"]
            elif total_errores <= self.total_folios * 5:
                estado = "OBSERVADO"
                detalles = [f"Se detectaron {total_errores} errores ortográficos en el documento"]
            else:
                estado = "RECHAZADO"
                detalles = [f"Cantidad elevada de errores ortográficos: {total_errores}"]
            
            for error in errores_por_folio[:10]:
                detalles.append(f"Folio {error['folio']}: {error['cantidad_errores']} errores")
            
            resultado = ResultadoVerificacion(
                tipo_verificacion="1.3 Verificación Ortográfica",
                estado=estado,
                detalles=detalles,
                porcentaje_cumplimiento=max(0, 100 - (total_errores / self.total_folios)),
                folios_afectados=[e['folio'] for e in errores_por_folio]
            )
            self.resultados.append(resultado)
            
        except Exception as e:
            print(f"Error en verificación ortográfica: {e}")
            resultado = ResultadoVerificacion(
                tipo_verificacion="1.3 Verificación Ortográfica",
                estado="NO PROCESADO",
                detalles=[f"No se pudo completar la verificación: {str(e)}"]
            )
            self.resultados.append(resultado)
    
    def _preparar_resultados(self) -> Dict:
        """Prepara el diccionario de resultados finales"""
        total_verificaciones = len(self.resultados)
        aprobados = sum(1 for r in self.resultados if r.estado == "APROBADO")
        observados = sum(1 for r in self.resultados if r.estado == "OBSERVADO")
        rechazados = sum(1 for r in self.resultados if r.estado == "RECHAZADO")
        
        return {
            'documento': self.ruta_documento,
            'fecha_verificacion': datetime.now().strftime(self.config['formato_fecha']),
            'total_folios': self.total_folios,
            'resumen': {
                'total_verificaciones': total_verificaciones,
                'aprobados': aprobados,
                'observados': observados,
                'rechazados': rechazados,
                'estado_global': self._determinar_estado_global()
            },
            'resultados': self.resultados
        }
    
    def _determinar_estado_global(self) -> str:
        """Determina el estado global del documento"""
        rechazados = sum(1 for r in self.resultados if r.estado == "RECHAZADO")
        observados = sum(1 for r in self.resultados if r.estado == "OBSERVADO")
        
        if rechazados > 0:
            return "NO ADMISIBLE"
        elif observados > 0:
            return "ADMISIBLE CON OBSERVACIONES"
        else:
            return "ADMISIBLE"
    
    def _generar_resultado_error(self) -> Dict:
        """Genera resultado en caso de error fatal"""
        return {
            'documento': self.ruta_documento,
            'fecha_verificacion': datetime.now().strftime(self.config['formato_fecha']),
            'error': 'No se pudo procesar el documento',
            'resultados': self.resultados
        }


class GeneradorReportePDF:
    """Generador de reportes PDF detallados"""
    
    def __init__(self, resultados: Dict, ruta_salida: str):
        self.resultados = resultados
        self.ruta_salida = ruta_salida
        self.estilos = getSampleStyleSheet()
        self._configurar_estilos()
    
    def _configurar_estilos(self):
        """Configura estilos personalizados para el reporte"""
        self.estilos.add(ParagraphStyle(
            name='TituloReporte',
            parent=self.estilos['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#1a1a1a'),
            spaceAfter=30,
            alignment=1
        ))
        
        self.estilos.add(ParagraphStyle(
            name='Seccion',
            parent=self.estilos['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=12,
            spaceBefore=12
        ))
    
    def generar(self):
        """Genera el reporte PDF completo"""
        print(f"\nGenerando reporte PDF: {self.ruta_salida}")
        
        doc = SimpleDocTemplate(
            self.ruta_salida,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        elementos = []
        elementos.extend(self._generar_encabezado())
        elementos.append(Spacer(1, 0.3 * inch))
        elementos.extend(self._generar_resumen_ejecutivo())
        elementos.append(Spacer(1, 0.3 * inch))
        elementos.extend(self._generar_resultados_detallados())
        elementos.append(PageBreak())
        elementos.extend(self._generar_conclusiones())
        
        doc.build(elementos)
        print(f"Reporte generado exitosamente: {self.ruta_salida}")
    
    def _generar_encabezado(self) -> List:
        """Genera el encabezado del reporte"""
        elementos = []
        
        titulo = Paragraph(
            "REPORTE DE VERIFICACIÓN DE ADMISIBILIDAD DOCUMENTAL",
            self.estilos['TituloReporte']
        )
        elementos.append(titulo)
        
        datos = [
            ['Documento analizado:', os.path.basename(self.resultados['documento'])],
            ['Fecha de verificación:', self.resultados['fecha_verificacion']],
            ['Total de folios:', str(self.resultados['total_folios'])],
            ['Estado global:', self.resultados['resumen']['estado_global']]
        ]
        
        tabla_info = Table(datos, colWidths=[2.5 * inch, 4 * inch])
        tabla_info.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#ecf0f1')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        
        elementos.append(tabla_info)
        return elementos
    
    def _generar_resumen_ejecutivo(self) -> List:
        """Genera el resumen ejecutivo"""
        elementos = []
        
        elementos.append(Paragraph("RESUMEN EJECUTIVO", self.estilos['Seccion']))
        
        resumen = self.resultados['resumen']
        
        datos = [
            ['CONCEPTO', 'CANTIDAD', 'PORCENTAJE'],
            ['Verificaciones realizadas', str(resumen['total_verificaciones']), '100%'],
            ['Aprobadas', str(resumen['aprobados']), 
             f"{round(resumen['aprobados']/resumen['total_verificaciones']*100, 1)}%"],
            ['Observadas', str(resumen['observados']), 
             f"{round(resumen['observados']/resumen['total_verificaciones']*100, 1)}%"],
            ['Rechazadas', str(resumen['rechazados']), 
             f"{round(resumen['rechazados']/resumen['total_verificaciones']*100, 1)}%"]
        ]
        
        tabla_resumen = Table(datos, colWidths=[3 * inch, 1.5 * inch, 1.5 * inch])
        tabla_resumen.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#34495e')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
        ]))
        
        elementos.append(tabla_resumen)
        elementos.append(Spacer(1, 0.2 * inch))
        
        estado_texto = f"<b>ESTADO GLOBAL: {resumen['estado_global']}</b>"
        elementos.append(Paragraph(estado_texto, self.estilos['Normal']))
        
        return elementos
    
    def _generar_resultados_detallados(self) -> List:
        """Genera la sección de resultados detallados"""
        elementos = []
        
        elementos.append(PageBreak())
        elementos.append(Paragraph("RESULTADOS DETALLADOS POR VERIFICACIÓN", self.estilos['Seccion']))
        elementos.append(Spacer(1, 0.2 * inch))
        
        for resultado in self.resultados['resultados']:
            elementos.extend(self._generar_seccion_resultado(resultado))
            elementos.append(Spacer(1, 0.3 * inch))
        
        return elementos
    
    def _generar_seccion_resultado(self, resultado: ResultadoVerificacion) -> List:
        """Genera una sección individual de resultado"""
        elementos = []
        
        titulo = f"<b>{resultado.tipo_verificacion}</b>"
        elementos.append(Paragraph(titulo, self.estilos['Normal']))
        elementos.append(Spacer(1, 0.1 * inch))
        
        color_estado = self._obtener_color_estado(resultado.estado)
        
        datos = [
            ['Estado:', resultado.estado],
            ['Cumplimiento:', f"{resultado.porcentaje_cumplimiento:.2f}%"],
            ['Folios afectados:', str(len(resultado.folios_afectados))]
        ]
        
        tabla_info = Table(datos, colWidths=[2 * inch, 4 * inch])
        tabla_info.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#ecf0f1')),
            ('BACKGROUND', (1, 0), (1, 0), color_estado),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey)
        ]))
        
        elementos.append(tabla_info)
        elementos.append(Spacer(1, 0.1 * inch))
        
        if resultado.detalles:
            elementos.append(Paragraph("<b>Detalles:</b>", self.estilos['Normal']))
            for detalle in resultado.detalles:
                texto_detalle = f"• {detalle}"
                elementos.append(Paragraph(texto_detalle, self.estilos['Normal']))
        
        if resultado.folios_afectados and len(resultado.folios_afectados) <= 20:
            folios_texto = ", ".join(map(str, resultado.folios_afectados))
            elementos.append(Paragraph(f"<b>Folios específicos:</b> {folios_texto}", self.estilos['Normal']))
        elif len(resultado.folios_afectados) > 20:
            primeros = ", ".join(map(str, resultado.folios_afectados[:20]))
            elementos.append(Paragraph(
                f"<b>Folios específicos (primeros 20):</b> {primeros} ... y {len(resultado.folios_afectados) - 20} más",
                self.estilos['Normal']
            ))
        
        return elementos
    
    def _generar_conclusiones(self) -> List:
        """Genera la sección de conclusiones y recomendaciones"""
        elementos = []
        
        elementos.append(Paragraph("CONCLUSIONES Y RECOMENDACIONES", self.estilos['Seccion']))
        elementos.append(Spacer(1, 0.2 * inch))
        
        estado_global = self.resultados['resumen']['estado_global']
        
        if estado_global == "ADMISIBLE":
            conclusion = """El documento ha superado satisfactoriamente todas las verificaciones de admisibilidad. 
            Se encuentra en condiciones óptimas para su procesamiento y archivo."""
        elif estado_global == "ADMISIBLE CON OBSERVACIONES":
            conclusion = """El documento presenta observaciones que, si bien no impiden su admisibilidad, 
            requieren atención y corrección para garantizar la calidad y consistencia documental."""
        else:
            conclusion = """El documento NO cumple con los requisitos mínimos de admisibilidad. 
            Se requiere corrección obligatoria de las deficiencias identificadas antes de su aceptación."""
        
        elementos.append(Paragraph(f"<b>Conclusión:</b> {conclusion}", self.estilos['Normal']))
        elementos.append(Spacer(1, 0.2 * inch))
        
        elementos.append(Paragraph("<b>Recomendaciones:</b>", self.estilos['Normal']))
        elementos.append(Spacer(1, 0.1 * inch))
        
        recomendaciones = self._generar_recomendaciones()
        for i, recom in enumerate(recomendaciones, 1):
            elementos.append(Paragraph(f"{i}. {recom}", self.estilos['Normal']))
        
        elementos.append(Spacer(1, 0.3 * inch))
        elementos.append(Paragraph("_" * 60, self.estilos['Normal']))
        elementos.append(Spacer(1, 0.1 * inch))
        elementos.append(Paragraph("Sistema Automatizado de Verificación de Admisibilidad", self.estilos['Normal']))
        elementos.append(Paragraph(f"Fecha de emisión: {self.resultados['fecha_verificacion']}", self.estilos['Normal']))
        
        return elementos
    
    def _generar_recomendaciones(self) -> List[str]:
        """Genera recomendaciones basadas en los resultados"""
        recomendaciones = []
        
        for resultado in self.resultados['resultados']:
            if resultado.estado == "RECHAZADO":
                if "Hojas en Blanco" in resultado.tipo_verificacion:
                    recomendaciones.append("Eliminar las hojas en blanco identificadas del documento")
                elif "Folios Duplicados" in resultado.tipo_verificacion:
                    recomendaciones.append("Revisar y eliminar los folios duplicados detectados")
                elif "Foliación" in resultado.tipo_verificacion:
                    recomendaciones.append("Corregir la numeración correlativa de los folios según normativa vigente")
                elif "Ilegibilidad" in resultado.tipo_verificacion:
                    recomendaciones.append("Re-escanear o reemplazar los folios con baja legibilidad")
                elif "Ortográfica" in resultado.tipo_verificacion:
                    recomendaciones.append("Realizar corrección ortográfica exhaustiva del documento")
            
            elif resultado.estado == "OBSERVADO":
                if "Plagio" in resultado.tipo_verificacion:
                    recomendaciones.append("Revisar las secciones con alta similitud para verificar autoría")
                elif "Ilegibilidad" in resultado.tipo_verificacion:
                    recomendaciones.append("Considerar mejorar la calidad de los folios observados")
        
        if not recomendaciones:
            recomendaciones.append("El documento cumple satisfactoriamente con todos los requisitos")
            recomendaciones.append("Proceder con el siguiente paso del proceso de admisión")
        
        return recomendaciones
    
    def _obtener_color_estado(self, estado: str):
        """Retorna el color correspondiente al estado"""
        colores = {
            'APROBADO': colors.HexColor('#27ae60'),
            'OBSERVADO': colors.HexColor('#f39c12'),
            'RECHAZADO': colors.HexColor('#e74c3c'),
            'NO PROCESADO': colors.HexColor('#95a5a6'),
            'ADMISIBLE': colors.HexColor('#27ae60'),
            'ADMISIBLE CON OBSERVACIONES': colors.HexColor('#f39c12'),
            'NO ADMISIBLE': colors.HexColor('#e74c3c')
        }
        return colores.get(estado, colors.grey)


def ejecutar_verificacion_rpa(ruta_documento: str, ruta_salida_reporte: str = None):
    """Función principal para ejecutar el RPA de verificación"""
    print("=" * 80)
    print("SISTEMA RPA DE VERIFICACIÓN DE ADMISIBILIDAD DOCUMENTAL")
    print("=" * 80)
    
    if not os.path.exists(ruta_documento):
        print(f"ERROR: El archivo {ruta_documento} no existe")
        return None
    
    configuracion = {
        'umbral_blanco': 0.98,
        'umbral_ilegibilidad': 0.60,
        'umbral_plagio': 0.85,
        'verificar_ortografia': True,
        'idioma_ortografia': 'es',
        'formato_fecha': '%Y-%m-%d %H:%M:%S'
    }
    
    verificador = VerificadorAdmisibilidad(ruta_documento, configuracion)
    resultados = verificador.ejecutar_verificacion_completa()
    
    if ruta_salida_reporte is None:
        nombre_base = Path(ruta_documento).stem
        ruta_salida_reporte = f"reporte_admisibilidad_{nombre_base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    
    generador = GeneradorReportePDF(resultados, ruta_salida_reporte)
    generador.generar()
    
    print("\n" + "=" * 80)
    print("RESUMEN DE VERIFICACIÓN")
    print("=" * 80)
    print(f"Estado global: {resultados['resumen']['estado_global']}")
    print(f"Verificaciones aprobadas: {resultados['resumen']['aprobados']}")
    print(f"Verificaciones observadas: {resultados['resumen']['observados']}")
    print(f"Verificaciones rechazadas: {resultados['resumen']['rechazados']}")
    print(f"\nReporte detallado guardado en: {ruta_salida_reporte}")
    print("=" * 80)
    
    return resultados


if __name__ == "__main__":
    RUTA_DOCUMENTO = "documento_ejemplo.pdf"
    RUTA_REPORTE = "reporte_admisibilidad.pdf"
    
    try:
        resultados = ejecutar_verificacion_rpa(RUTA_DOCUMENTO, RUTA_REPORTE)
        
        if resultados:
            print("\nVerificación completada exitosamente")
            
            ruta_json = RUTA_REPORTE.replace('.pdf', '.json')
            with open(ruta_json, 'w', encoding='utf-8') as f:
                json.dump({
                    'documento': resultados['documento'],
                    'fecha_verificacion': resultados['fecha_verificacion'],
                    'total_folios': resultados['total_folios'],
                    'resumen': resultados['resumen'],
                    'resultados': [
                        {
                            'tipo': r.tipo_verificacion,
                            'estado': r.estado,
                            'detalles': r.detalles,
                            'porcentaje': r.porcentaje_cumplimiento,
                            'folios_afectados': r.folios_afectados
                        } for r in resultados['resultados']
                    ]
                }, f, indent=2, ensure_ascii=False)
            
            print(f"Resultados también guardados en formato JSON: {ruta_json}")
    
    except Exception as e:
        print(f"Error durante la ejecución: {e}")
        import traceback
        traceback.print_exc()
