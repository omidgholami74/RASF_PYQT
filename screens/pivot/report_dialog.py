from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QTextEdit
import re

class ReportDialog(QDialog):
    """Dialog to display both text-based and table-based CRM analysis report."""
    def __init__(self, parent, annotations):
        super().__init__(parent)
        self.annotations = annotations
        self.setWindowTitle("Professional CRM Analysis Report")
        self.setGeometry(200, 200, 900, 700)
        layout = QVBoxLayout(self)
        
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

    def generate_html_report(self):
        html = """
        <html>
        <head>
            <style>
                body { font-family: 'Segoe UI', sans-serif; color: #333; line-height: 1.6; }
                h1 { color: #007bff; text-align: center; margin-bottom: 20px; }
                h2 { color: #0056b3; margin-top: 30px; }
                .report-section { margin-bottom: 30px; border-bottom: 1px solid #dee2e6; padding-bottom: 10px; }
                .crm-title { font-weight: bold; color: #28a745; }
                .problematic { color: #dc3545; font-weight: bold; }
                .in-range { color: #28a745; }
                .out-range { color: #dc3545; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                th, td { border: 1px solid #dee2e6; padding: 8px; text-align: left; }
                th { background-color: #007bff; color: white; }
            </style>
        </head>
        <body>
            <h1>CRM Analysis Report</h1>
            <h2>Text Report</h2>
        """
        
        # Text-based report
        for annotation in self.annotations:
            lines = annotation.split('\n')
            crm_header = lines[0].strip()
            details = lines[1:] if len(lines) > 1 else []
            
            html += f'<div class="report-section">'
            html += f'<span class="crm-title">{crm_header}</span><br>'
            
            for detail in details:
                detail = detail.strip()
                if not detail:
                    continue
                if "in range" in detail:
                    html += f'<span class="in-range">{detail}</span><br>'
                elif "out of range" in detail:
                    html += f'<span class="out-range">{detail}</span><br>'
                elif "Scaling needed" in detail:
                    if "problematic" in detail:
                        html += f'<span class="problematic">{detail}</span><br>'
                    else:
                        html += f'<span>{detail}</span><br>'
                else:
                    html += f'{detail}<br>'
            
            html += '</div>'
        
        # Table-based report
        html += """
            <h2>Table Report</h2>
            <table>
                <tr>
                    <th>Verification ID</th>
                    <th>Check Verification Value</th>
                    <th>Pivot Verification Value</th>
                    <th>Original Range</th>
                    <th>Status</th>
                    <th>Blank Value Subtracted</th>
                    <th>Corrected Verification Value</th>
                    <th>Corrected Range</th>
                    <th>Status after Blank Subtraction</th>
                    <th>Required Scaling (%)</th>
                </tr>
        """
        
        for annotation in self.annotations:
            lines = annotation.split('\n')
            if not lines:
                continue
            
            verification_id_match = re.match(r'Verification ID: (\w+) \(Label: .+\)', lines[0].strip())
            verification_id = verification_id_match.group(1) if verification_id_match else "Unknown"
            
            check_val = pivot_val = orig_range = status = blank_val = corrected_val = corrected_range = corrected_status = scaling = ""
            
            for line in lines[1:]:
                line = line.strip()
                if not line:
                    continue
                if "Check Verification Value:" in line:
                    check_val = line.split(": ")[1].strip()
                elif "Pivot Verification Value:" in line:
                    pivot_val = line.split(": ")[1].strip()
                elif "Original Range:" in line:
                    orig_range = line.split(": ")[1].strip()
                elif "Status: In range" in line:
                    status = '<span class="in-range">In range</span>'
                elif "Status: Out of range" in line:
                    status = '<span class="out-range">Out of range</span>'
                elif "Blank Value Subtracted:" in line:
                    blank_val = line.split(": ")[1].strip()
                elif "Corrected Verification Value:" in line:
                    corrected_val = line.split(": ")[1].strip()
                elif "Corrected Range:" in line:
                    corrected_range = line.split(": ")[1].strip()
                elif "Status after Blank Subtraction: In range" in line:
                    corrected_status = '<span class="in-range">In range</span>'
                elif "Status after Blank Subtraction: Out of range" in line:
                    corrected_status = '<span class="out-range">Out of range</span>'
                elif "Required Scaling:" in line:
                    scaling_match = re.search(r'Required Scaling: ([\d.]+)% (increase|decrease)', line)
                    if scaling_match:
                        scaling = f"{scaling_match.group(1)}% {scaling_match.group(2)}"
                        if "Scaling exceeds 200%" in line:
                            scaling = f'<span class="problematic">{scaling} (Problematic)</span>'
            
            html += f"""
                <tr>
                    <td>{verification_id}</td>
                    <td>{check_val}</td>
                    <td>{pivot_val}</td>
                    <td>{orig_range}</td>
                    <td>{status}</td>
                    <td>{blank_val}</td>
                    <td>{corrected_val}</td>
                    <td>{corrected_range}</td>
                    <td>{corrected_status}</td>
                    <td>{scaling}</td>
                </tr>
            """
        
        html += """
            </table>
        </body>
        </html>
        """
        self.text_edit.setHtml(html)