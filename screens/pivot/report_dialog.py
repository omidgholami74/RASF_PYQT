from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit, QCheckBox, QScrollArea, QWidget, QLabel
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtCore import Qt
import re

class ReportDialog(QDialog):
    """Dialog to display table-based CRM analysis report with scrollable column visibility toggles."""
    def __init__(self, parent, annotations):
        super().__init__(parent)
        self.annotations = annotations
        self.setWindowTitle("Professional CRM Analysis Report")

        # Set window to full width
        screen = QGuiApplication.primaryScreen().size()
        self.setGeometry(0, 200, screen.width(), 700)
        
        # Initialize column visibility dictionary (not all checked by default)
        self.column_visibility = {
            'Verification ID': True,
            'Certificate Value': True,
            'Sample Value': True,
            'Acceptable Range': True,
            'Blank Value': False,
            'Blank Correction Status': False,
            'Sample Value - Blank': False,
            'Soln Conc': True,
            'Int': False,
            'Calibration Range': True,
            'ICP Recovery (%)': False,
            'ICP Status': False,
            'ICP Detection Limit': False,
            'ICP RSD%': False,
            'Required Scaling (%)': False
        }
        
        layout = QVBoxLayout(self)
        
        # Add scrollable checkbox area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setFixedHeight(50)  # Adjust height to fit checkboxes
        
        checkbox_widget = QWidget()
        checkbox_layout = QHBoxLayout(checkbox_widget)
        checkbox_layout.addWidget(QLabel("Show Columns:"))
        self.checkboxes = {}
        for column in self.column_visibility.keys():
            checkbox = QCheckBox(column)
            checkbox.setChecked(self.column_visibility[column])
            checkbox.toggled.connect(self.update_report)
            self.checkboxes[column] = checkbox
            checkbox_layout.addWidget(checkbox)
        checkbox_layout.addStretch()
        
        scroll_area.setWidget(checkbox_widget)
        layout.addWidget(scroll_area)
        
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
                padding: 10px;
            }
        """)
        layout.addWidget(self.text_edit)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        layout.addWidget(close_btn)
        
        self.generate_html_report()

    def update_report(self):
        """Update column visibility and regenerate the report."""
        for column, checkbox in self.checkboxes.items():
            self.column_visibility[column] = checkbox.isChecked()
        self.generate_html_report()

    def generate_html_report(self):
        """Generate HTML report with only visible columns."""
        html = """
        <html>
        <head>
            <style>
                body { font-family: 'Segoe UI', sans-serif; color: #333; line-height: 1.6; }
                h1 { color: #007bff; text-align: center; margin-bottom: 20px; }
                h2 { color: #0056b3; margin-top: 30px; }
                .problematic { color: #dc3545; font-weight: bold; }
                .in-range { background-color: #d4edda; }
                .out-range { background-color: #f8d7da; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                th, td { border: 1px solid #dee2e6; padding: 8px; text-align: left; color: #333; }
                th { background-color: #007bff; color: white; }
            </style>
        </head>
        <body>
            <h1>CRM Analysis Report</h1>
            <h2>Table Report</h2>
            <table>
                <tr>
        """
        
        # Add table headers for visible columns only
        for column in self.column_visibility:
            if self.column_visibility[column]:
                html += f"<th>{column}</th>"
        html += "</tr>"

        for annotation in self.annotations:
            lines = annotation.split('\n')
            if not lines:
                continue
            
            verification_id_match = re.match(r'Verification ID: (\w+) \(Label: .+\)', lines[0].strip())
            verification_id = verification_id_match.group(1) if verification_id_match else "Unknown"
            
            # Initialize variables for all possible columns
            certificate_val = sample_val = acceptable_range = blank_val = blank_correction_status = corrected_sample_val = soln_conc = int_val = calibration_range = icp_recovery = icp_status = detection_limit = rsd_percent = scaling = ""
            sample_val_class = corrected_sample_val_class = soln_conc_class = icp_status_class = ""
            
            # Parse annotation lines
            for line in lines[1:]:
                line = line.strip()
                if not line:
                    continue
                if "Certificate Value:" in line:
                    certificate_val = line.split(": ")[1].strip()
                elif "Sample Value:" in line:
                    sample_val = line.split(": ")[1].strip()
                elif "Acceptable Range:" in line:
                    acceptable_range = line.split(": ")[1].strip()
                elif "Status: In range" in line:
                    sample_val_class = 'in-range'
                elif "Status: Out of range" in line:
                    sample_val_class = 'out-range'
                elif "Blank Value:" in line:
                    blank_val = line.split(": ")[1].strip()
                elif "Blank Correction Status:" in line:
                    blank_correction_status = line.split(": ")[1].strip()
                elif "Sample Value - Blank:" in line:
                    corrected_sample_val = line.split(": ")[1].strip()
                    # Verify if Sample Value - Blank is within Acceptable Range
                    try:
                        corrected_val = float(corrected_sample_val)
                        range_match = re.match(r'\[([\d.]+) to ([\d.]+)\]', acceptable_range)
                        if range_match:
                            lower, upper = float(range_match.group(1)), float(range_match.group(2))
                            corrected_sample_val_class = 'in-range' if lower <= corrected_val <= upper else 'out-range'
                    except (ValueError, TypeError):
                        corrected_sample_val_class = 'out-range'
                elif "Soln Conc:" in line:
                    parts = line.split(": ", 1)[1].rsplit(" ", 1)
                    soln_conc = parts[0].strip()
                    soln_conc_status = parts[1].strip() if len(parts) > 1 else ""
                    soln_conc_class = 'in-range' if soln_conc_status == 'in_range' else 'out-range'
                elif "Int:" in line:
                    int_val = line.split(": ")[1].strip()
                elif "Calibration Range:" in line:
                    calibration_range = line.split(": ", 1)[1].rsplit(" ", 1)[0].strip()
                elif "ICP Recovery:" in line:
                    icp_recovery = line.split(": ", 1)[1].rsplit(" ", 1)[0].strip()
                elif "ICP Status:" in line:
                    icp_status = line.split(": ")[1].strip()
                    icp_status_class = 'in-range' if icp_status == 'In Range' else 'out-range'
                elif "ICP Detection Limit:" in line:
                    detection_limit = line.split(": ")[1].strip()
                elif "ICP RSD%:" in line:
                    rsd_percent = line.split(": ")[1].strip()
                elif "Required Scaling:" in line:
                    scaling_match = re.search(r'Required Scaling: ([\d.]+)% (increase|decrease)', line)
                    if scaling_match:
                        scaling = f"{scaling_match.group(1)}% {scaling_match.group(2)}"
                        if "Scaling exceeds 200%" in line:
                            scaling = f'<span class="problematic">{scaling} (Problematic)</span>'
            
            # Map column names to their values and classes
            column_data = {
                'Verification ID': (verification_id, ''),
                'Certificate Value': (certificate_val, ''),
                'Sample Value': (sample_val, sample_val_class),
                'Acceptable Range': (acceptable_range, ''),
                'Blank Value': (blank_val, ''),
                'Blank Correction Status': (blank_correction_status, ''),
                'Sample Value - Blank': (corrected_sample_val, corrected_sample_val_class),
                'Soln Conc': (soln_conc, soln_conc_class),
                'Int': (int_val, ''),
                'Calibration Range': (calibration_range, ''),
                'ICP Recovery (%)': (icp_recovery, ''),
                'ICP Status': (icp_status, icp_status_class),
                'ICP Detection Limit': (detection_limit, ''),
                'ICP RSD%': (rsd_percent, ''),
                'Required Scaling (%)': (scaling, '')
            }
            
            # Generate table row with only visible columns
            html += "<tr>"
            for column, (value, css_class) in column_data.items():
                if self.column_visibility[column]:
                    html += f'<td class="{css_class}">{value}</td>'
            html += "</tr>"
        
        html += """
            </table>
        </body>
        </html>
        """
        self.text_edit.setHtml(html)