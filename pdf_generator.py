"""
PDF Report Generator Module
Creates elegant executive PDF reports with matplotlib charts and AI conclusions.
"""

import os
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional
from io import BytesIO

# PDF Generation
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch, cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether, Flowable
)
from reportlab.pdfgen import canvas

# Charts
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)

# Spanish month names
MESES_ESP = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
             'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

# Professional color palette
COLORS = {
    'primary': '#1F497D',      # Professional blue
    'secondary': '#4472C4',    # Lighter blue
    'accent': '#2E75B6',       # Medium blue
    'text': '#333333',         # Dark gray text
    'light_gray': '#F2F2F2',   # Background gray
    'border': '#D9D9D9',       # Border gray
    'success': '#70AD47',      # Green for positive
    'warning': '#FFC000',      # Yellow for caution
}


class VerticalBar(Flowable):
    """Custom vertical bar separator."""
    
    def __init__(self, width=3, height=20, color=COLORS['primary']):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.color = HexColor(color)
    
    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


class PDFReportGenerator:
    """Generates professional PDF reports for Zabbix metrics."""
    
    def __init__(self, output_dir: str):
        """
        Initialize the PDF generator.
        
        Args:
            output_dir: Directory to save the PDF report
        """
        self.output_dir = output_dir
        self.styles = self._create_styles()
        self.items_data = []  # Store all items data for the report
        
        # Report configuration with defaults
        self.report_config = {}
        self.report_defaults = {
            'incidentes': "No se presentan incidentes de servicio.",
            'riesgos': "No se registran riesgos del servicio durante el periodo.",
            'alertas': "No se evidencian alertas que afecten la continuidad operativa.",
            'uptime_fecha': "",
            'uptime_servidor': "",
            'uptime_bd': "",
            'dim_rendimiento': "Sin observaciones",
            'dim_contingencia': "Sin observaciones",
            'dim_soporte': "Sin observaciones",
            'dim_actualizaciones': "Sin observaciones",
            'dim_respaldos': "Sin observaciones",
        }
        
        # Per-host configurations (incidentes, riesgos, alertas, uptime)
        self.host_configs = {}  # host_name -> {incidentes, riesgos, alertas, uptime_*}
        self.current_host = ""  # Track current host being generated
        
        # Storage/filesystem data per host
        self.storage_data = {}  # host_name -> [{'fsname', 'pused', 'used_gb', 'total_gb'}, ...]
    
    def set_report_config(self, config: Dict, defaults: Dict = None):
        """Set the report configuration from GUI inputs."""
        self.report_config = config
        if defaults:
            self.report_defaults = defaults
    
    def set_host_configs(self, host_configs: Dict):
        """Set per-host configurations."""
        self.host_configs = host_configs
    
    def _get_config_value(self, key: str) -> str:
        """Get config value, falling back to default if empty."""
        value = self.report_config.get(key, '').strip()
        if not value:
            return self.report_defaults.get(key, '')
        return value
    
    def _get_host_config_value(self, key: str) -> str:
        """Get per-host config value (incidentes, riesgos, alertas)."""
        host_config = self.host_configs.get(self.current_host, {})
        value = host_config.get(key, '').strip()
        if not value:
            return self.report_defaults.get(key, '')
        return value
    
    def _create_styles(self):
        """Create custom paragraph styles."""
        styles = getSampleStyleSheet()
        
        # Host header style
        styles.add(ParagraphStyle(
            name='HostHeader',
            fontName='Helvetica-Bold',
            fontSize=24,
            textColor=HexColor(COLORS['primary']),
            spaceAfter=10,
            alignment=TA_LEFT
        ))
        
        # Item title style (with vertical bar effect)
        styles.add(ParagraphStyle(
            name='ItemTitle',
            fontName='Helvetica-Bold',
            fontSize=14,
            textColor=HexColor(COLORS['primary']),
            spaceBefore=20,
            spaceAfter=10,
            leftIndent=10,
            borderPadding=5,
            borderColor=HexColor(COLORS['primary']),
            borderWidth=0,
        ))
        
        # Analysis paragraph style
        styles.add(ParagraphStyle(
            name='Analysis',
            fontName='Helvetica',
            fontSize=10,
            textColor=HexColor(COLORS['text']),
            alignment=TA_JUSTIFY,
            spaceBefore=10,
            spaceAfter=15,
            leading=14,
        ))
        
        # Data card label
        styles.add(ParagraphStyle(
            name='CardLabel',
            fontName='Helvetica',
            fontSize=9,
            textColor=HexColor('#666666'),
            alignment=TA_CENTER,
        ))
        
        # Data card value
        styles.add(ParagraphStyle(
            name='CardValue',
            fontName='Helvetica-Bold',
            fontSize=16,
            textColor=HexColor(COLORS['primary']),
            alignment=TA_CENTER,
        ))
        
        return styles
    
    def add_item_data(self, host_name: str, item_name: str, trends: List[Dict],
                      stats: Dict, conclusion: Optional[str] = None):
        """
        Add item data for the report.
        
        Args:
            host_name: Name of the host
            item_name: Name of the item
            trends: List of trend data points
            stats: Calculated statistics
            conclusion: AI-generated conclusion
        """
        self.items_data.append({
            'host': host_name,
            'item': item_name,
            'trends': trends,
            'stats': stats,
            'conclusion': conclusion
        })
        logger.info(f"Added item data: {host_name}/{item_name}")
    
    def _create_chart(self, trends: List[Dict], item_name: str, 
                      width: float = 6.5, height: float = 2.5) -> Image:
        """
        Create a modern line chart from trend data.
        
        Args:
            trends: List of trend data points
            item_name: Name of the item for the title
            width: Chart width in inches
            height: Chart height in inches
            
        Returns:
            ReportLab Image object
        """
        # Convert trends to DataFrame
        df = pd.DataFrame(trends)
        df['datetime'] = pd.to_datetime(df['clock'].astype(int), unit='s')
        df['value_avg'] = pd.to_numeric(df['value_avg'], errors='coerce')
        df['value_max'] = pd.to_numeric(df['value_max'], errors='coerce')
        df['value_min'] = pd.to_numeric(df['value_min'], errors='coerce')
        
        # Sort by datetime
        df = df.sort_values('datetime')
        
        # Create figure with modern styling
        fig, ax = plt.subplots(figsize=(width, height), dpi=100)
        
        # Remove top and right spines
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#CCCCCC')
        ax.spines['bottom'].set_color('#CCCCCC')
        
        # Plot lines
        ax.fill_between(df['datetime'], df['value_min'], df['value_max'], 
                       alpha=0.1, color=COLORS['secondary'], label='Rango')
        ax.plot(df['datetime'], df['value_avg'], 
               color=COLORS['primary'], linewidth=1.5, label='Promedio')
        
        # Subtle grid
        ax.grid(True, axis='y', linestyle='-', alpha=0.3, color='#E0E0E0')
        ax.set_axisbelow(True)
        
        # Format x-axis (day/month only, no hours)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(df)//8)))
        plt.xticks(fontsize=8, color='#666666')
        plt.yticks(fontsize=8, color='#666666')
        
        # Y-axis label
        ax.set_ylabel('%', fontsize=9, color='#666666')
        
        # Set y-axis limits with padding
        y_min = max(0, df['value_min'].min() - 5)
        y_max = min(100, df['value_max'].max() + 10)
        ax.set_ylim(y_min, y_max)
        
        # Tight layout
        plt.tight_layout()
        
        # Save to bytes buffer
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        plt.close(fig)
        buf.seek(0)
        
        # Create ReportLab Image
        img = Image(buf, width=width*inch, height=height*inch)
        return img
    
    def _create_data_cards(self, stats: Dict) -> Table:
        """
        Create data cards showing key metrics.
        
        Args:
            stats: Statistics dictionary
            
        Returns:
            ReportLab Table with data cards
        """
        # Card data
        cards = [
            ('PROMEDIO', f"{stats.get('avg_monthly', 'N/A')}%"),
            ('MÍNIMO', f"{stats.get('min_absolute', 'N/A')}%"),
            ('MÁXIMO', f"{stats.get('max_absolute', 'N/A')}%"),
            ('P95', f"{stats.get('p95', 'N/A')}%"),
        ]
        
        # Create table data
        card_row = []
        for label, value in cards:
            # Each card is a mini-table
            card_content = [
                [Paragraph(value, self.styles['CardValue'])],
                [Paragraph(label, self.styles['CardLabel'])]
            ]
            card_table = Table(card_content, colWidths=[1.4*inch])
            card_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), HexColor(COLORS['light_gray'])),
                ('BOX', (0, 0), (-1, -1), 1, HexColor(COLORS['border'])),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            card_row.append(card_table)
        
        # Main table containing all cards
        main_table = Table([card_row], colWidths=[1.5*inch]*4)
        main_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ]))
        
        return main_table
    
    def _create_item_title(self, item_name: str) -> Table:
        """
        Create item title with vertical bar indicator.
        
        Args:
            item_name: Name of the item
            
        Returns:
            Table with vertical bar and title
        """
        # Vertical bar cell
        bar_cell = Table([['']], colWidths=[4], rowHeights=[25])
        bar_cell.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), HexColor(COLORS['primary'])),
        ]))
        
        # Title
        title = Paragraph(item_name.upper(), self.styles['ItemTitle'])
        
        # Combine - wider title column to prevent cutoff
        title_table = Table([[bar_cell, title]], colWidths=[8, 6.5*inch])
        title_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (0, 0), 0),
            ('LEFTPADDING', (1, 0), (1, 0), 10),
        ]))
        
        return title_table
    
    def _create_divider(self) -> Table:
        """Create an elegant horizontal divider."""
        divider = Table([['']], colWidths=[6.5*inch], rowHeights=[2])
        divider.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), HexColor(COLORS['border'])),
            ('TOPPADDING', (0, 0), (0, 0), 0),
            ('BOTTOMPADDING', (0, 0), (0, 0), 0),
        ]))
        return divider
    def _create_operative_block(self) -> List:
        """Create Section A: Operative Text Block with incidents, risks, alerts."""
        elements = []
        
        # Intro paragraph
        intro = ("A continuación, se presentan los principales tópicos operativos del servicio, "
                 "incluyendo incidentes, riesgos y alertas relevantes del periodo.")
        elements.append(Paragraph(intro, self.styles['Analysis']))
        elements.append(Spacer(1, 10))
        
        # Compact format: using per-host config values
        incidentes = self._get_host_config_value('incidentes')
        riesgos = self._get_host_config_value('riesgos')
        alertas = self._get_host_config_value('alertas')
        
        elements.append(Paragraph(f"<b>Incidentes del Servicio:</b> {incidentes}", self.styles['Analysis']))
        elements.append(Spacer(1, 5))
        
        elements.append(Paragraph(f"<b>Riesgos del Servicio:</b> {riesgos}", self.styles['Analysis']))
        elements.append(Spacer(1, 5))
        
        elements.append(Paragraph(f"<b>Incidentes de Alerta:</b> {alertas}", self.styles['Analysis']))
        elements.append(Spacer(1, 15))
        
        return elements
    
    def _create_severity_glossary(self) -> List:
        """Create Section B: Severity Glossary (G, M, B)."""
        elements = []
        
        elements.append(Paragraph("<b>Glosario de Gravedad</b>", self.styles['ItemTitle']))
        elements.append(Spacer(1, 10))
        
        # Glossary table
        data = [
            ['Símbolo', 'Nivel', 'Descripción'],
            ['G', 'Grave', 'Requiere atención inmediata'],
            ['M', 'Medio', 'Requiere seguimiento'],
            ['B', 'Bajo', 'Informativo'],
        ]
        
        table = Table(data, colWidths=[0.8*inch, 1*inch, 3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(COLORS['primary'])),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor(COLORS['border'])),
            ('BACKGROUND', (0, 1), (-1, -1), HexColor(COLORS['light_gray'])),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, HexColor(COLORS['light_gray'])]),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        return elements
    
    def _create_dimensions_table(self) -> List:
        """Create Section C: Dimension Summary Table."""
        elements = []
        
        elements.append(Paragraph("<b>Tabla Resumen de Dimensiones</b>", self.styles['ItemTitle']))
        elements.append(Spacer(1, 10))
        
        # Dimensions table
        data = [
            ['Dimensión', 'Estado'],
            ['Rendimiento', self._get_config_value('dim_rendimiento')],
            ['Contingencia', self._get_config_value('dim_contingencia')],
            ['Soporte', self._get_config_value('dim_soporte')],
            ['Actualizaciones', self._get_config_value('dim_actualizaciones')],
            ['Respaldos', self._get_config_value('dim_respaldos')],
        ]
        
        table = Table(data, colWidths=[1.5*inch, 4.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(COLORS['primary'])),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor(COLORS['border'])),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, HexColor(COLORS['light_gray'])]),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        return elements
    
    def _create_uptime_section(self) -> List:
        """Create Section D: Uptime Information (per-host)."""
        elements = []
        
        # Get per-host uptime values
        uptime_fecha = self._get_host_config_value('uptime_fecha')
        uptime_servidor = self._get_host_config_value('uptime_servidor')
        uptime_bd = self._get_host_config_value('uptime_bd')
        
        # Check if empty (not just default)
        host_config = self.host_configs.get(self.current_host, {})
        if not (host_config.get('uptime_fecha') or host_config.get('uptime_servidor') or host_config.get('uptime_bd')):
            return elements
        
        elements.append(Paragraph("<b>Información de Uptime</b>", self.styles['ItemTitle']))
        elements.append(Spacer(1, 10))
        
        # Uptime data
        data = [['Métrica', 'Valor']]
        if uptime_fecha:
            data.append(['Fecha Último Inicio', uptime_fecha])
        if uptime_servidor:
            data.append(['Uptime Servidor', uptime_servidor])
        if uptime_bd:
            data.append(['Uptime BD', uptime_bd])
        
        table = Table(data, colWidths=[2*inch, 4*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(COLORS['secondary'])),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor(COLORS['border'])),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        return elements
    
    def add_storage_data(self, host_name: str, fs_data: List[Dict]):
        """
        Add filesystem/storage data for a host.
        
        Args:
            host_name: Name of the host
            fs_data: List of filesystem dicts [{fsname, pused, used_gb, total_gb}, ...]
        """
        if fs_data:
            self.storage_data[host_name] = fs_data
            logger.info(f"Added storage data for {host_name}: {len(fs_data)} partitions")
    
    def _create_storage_section(self) -> List:
        """Create storage/partitions summary table for current host."""
        elements = []
        
        # Get storage data for current host
        fs_data = self.storage_data.get(self.current_host, [])
        if not fs_data:
            return elements
        
        # Section title
        elements.append(Spacer(1, 15))
        elements.append(Paragraph("💾 ALMACENAMIENTO", self.styles['SectionHeader']))
        elements.append(Spacer(1, 8))
        
        # Table header
        header = ['Partición', '% Usado', 'Usado (GB)', 'Total (GB)', 'Estado']
        data = [header]
        
        # Add rows for each filesystem
        for fs in fs_data:
            fsname = fs.get('fsname', 'N/A')
            pused = fs.get('pused', 0)
            used_gb = fs.get('used_gb', 0)
            total_gb = fs.get('total_gb', 0)
            
            # Status indicator based on usage
            if pused >= 90:
                status = '🔴 Crítico'
            elif pused >= 80:
                status = '🟠 Alto'
            elif pused >= 70:
                status = '🟡 Medio'
            else:
                status = '🟢 Normal'
            
            data.append([
                fsname,
                f"{pused:.1f}%",
                f"{used_gb:.2f}",
                f"{total_gb:.2f}",
                status
            ])
        
        # Create table
        col_widths = [2.5*inch, 0.9*inch, 1*inch, 1*inch, 1*inch]
        table = Table(data, colWidths=col_widths)
        
        # Style the table with color coding
        style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(COLORS['primary'])),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor(COLORS['border'])),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]
        
        # Add row colors based on usage
        for i, fs in enumerate(fs_data, 1):
            pused = fs.get('pused', 0)
            if pused >= 90:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), HexColor('#FFEBEE')))  # Light red
            elif pused >= 80:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), HexColor('#FFF3E0')))  # Light orange
            elif pused >= 70:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), HexColor('#FFFDE7')))  # Light yellow
            else:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), white))
        
        table.setStyle(TableStyle(style_commands))
        elements.append(table)
        elements.append(Spacer(1, 15))
        
        return elements
    
    def generate_report(self, report_name: str = "informe_ejecutivo") -> str:
        """
        Generate the PDF report with all added items.
        
        Args:
            report_name: Base name for the PDF file
            
        Returns:
            Path to the generated PDF file
        """
        if not self.items_data:
            logger.warning("No items data to generate report")
            return ""
        
        # Create PDF path
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        pdf_path = os.path.join(self.output_dir, f"{report_name}_{timestamp}.pdf")
        
        # Create document with proper margins
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=letter,
            leftMargin=0.5*inch,
            rightMargin=0.5*inch,
            topMargin=0.5*inch,
            bottomMargin=0.5*inch
        )
        
        # Build story (content)
        story = []
        
        # Group items by host
        hosts = {}
        for item_data in self.items_data:
            host = item_data['host']
            if host not in hosts:
                hosts[host] = []
            hosts[host].append(item_data)
        
        # Generate content for each host
        for host_idx, (host_name, items) in enumerate(hosts.items()):
            # Page break between hosts (except first)
            if host_idx > 0:
                story.append(PageBreak())
            
            # Set current host for per-host config values
            self.current_host = host_name
            
            # Host header
            story.append(Paragraph(host_name, self.styles['HostHeader']))
            story.append(Spacer(1, 10))  # 10px space before divider
            story.append(self._create_divider())
            story.append(Spacer(1, 20))
            
            # Date info in Spanish (day month year, no hours)
            now = datetime.now()
            mes_esp = MESES_ESP[now.month - 1]
            date_info = f"Reporte generado: {now.day} de {mes_esp} de {now.year}"
            story.append(Paragraph(date_info, self.styles['Analysis']))
            story.append(Spacer(1, 15))
            
            # === Section A: Operative Block ===
            story.extend(self._create_operative_block())
            
            # === Section B: Severity Glossary ===
            story.extend(self._create_severity_glossary())
            
            # === Section C: Dimensions Table ===
            story.extend(self._create_dimensions_table())
            
            # === Section D: Uptime Section ===
            story.extend(self._create_uptime_section())
            
            # === Section E: Dynamic Items (Zabbix) ===
            story.append(Paragraph("<b>Ítems de Monitoreo</b>", self.styles['ItemTitle']))
            story.append(Spacer(1, 15))
            
            # Each item for this host
            for item_data in items:
                item_section = []
                
                # Item title with vertical bar
                item_section.append(self._create_item_title(item_data['item']))
                item_section.append(Spacer(1, 10))
                
                # Chart
                if item_data['trends']:
                    try:
                        chart = self._create_chart(item_data['trends'], item_data['item'])
                        item_section.append(chart)
                        item_section.append(Spacer(1, 10))
                    except Exception as e:
                        logger.error(f"Failed to create chart: {str(e)}")
                
                # Data cards
                if item_data['stats']:
                    cards = self._create_data_cards(item_data['stats'])
                    item_section.append(cards)
                    item_section.append(Spacer(1, 15))
                
                # AI Analysis paragraph
                if item_data['conclusion']:
                    analysis_text = item_data['conclusion'].replace('\n', ' ').strip()
                    story.append(KeepTogether(item_section))
                    story.append(Paragraph(
                        f"<b>Análisis:</b> {analysis_text}",
                        self.styles['Analysis']
                    ))
                else:
                    story.append(KeepTogether(item_section))
                
                story.append(Spacer(1, 25))
            
            # === Section F: Storage Table (if available) ===
            story.extend(self._create_storage_section())
        
        # Build PDF
        try:
            doc.build(story)
            logger.info(f"PDF report generated: {pdf_path}")
            return pdf_path
        except Exception as e:
            logger.error(f"Failed to generate PDF: {str(e)}")
            return ""
    
    def clear_data(self):
        """Clear all stored items data."""
        self.items_data = []
