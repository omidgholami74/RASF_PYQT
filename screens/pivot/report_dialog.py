from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QTextEdit

class ReportDialog(QDialog):
    """Dialog to display report with improved UI."""
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
                .report-section { margin-bottom: 30px; border-bottom: 1px solid #dee2e6; padding-bottom: 10px; }
                .crm-title { font-weight: bold; color: #28a745; }
                .problematic { color: #dc3545; font-weight: bold; }
                .in-range { color: #28a745; }
                .out-range { color: #dc3545; }
            </style>
        </head>
        <body>
            <h1>CRM Analysis Report</h1>
        """
        
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
        
        html += """
        </body>
        </html>
        """
        self.text_edit.setHtml(html)