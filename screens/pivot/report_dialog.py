from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit, QCheckBox, QScrollArea, QWidget, QLabel
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtCore import Qt
import re
import numpy as np
from scipy.optimize import differential_evolution

class ReportDialog(QDialog):
    """Dialog to display table-based CRM analysis report with scrollable column visibility toggles and textual decision analysis."""
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
        """Generate HTML report with only visible columns and add textual decision analysis section."""
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
                .decision-section { margin-top: 40px; padding: 15px; background-color: #e9ecef; border-radius: 8px; }
                .decision-section ul { list-style-type: disc; margin-left: 20px; }
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

        crm_data = []  # Collect for decision analysis
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
            
            # Collect for decision analysis (parse numbers where possible)
            try:
                crm_entry = {
                    'id': verification_id,
                    'cert_val': float(certificate_val) if certificate_val else None,
                    'sample_val': float(sample_val) if sample_val else None,
                    'lower': float(re.match(r'\[([\d.]+) to ([\d.]+)\]', acceptable_range).group(1)) if acceptable_range else None,
                    'upper': float(re.match(r'\[([\d.]+) to ([\d.]+)\]', acceptable_range).group(2)) if acceptable_range else None,
                    'blank_val': float(blank_val) if blank_val else 0,
                    'soln_conc': float(soln_conc) if soln_conc else None,
                    'wavelength': 'default'  # Extend if multiple
                }
                if all(v is not None for v in [crm_entry['cert_val'], crm_entry['sample_val'], crm_entry['lower'], crm_entry['upper']]):
                    crm_data.append(crm_entry)
            except:
                pass
        
        html += "</table>"
        
        # Add textual decision analysis section
        html += self.generate_decision_analysis(crm_data)
        
        html += """
        </body>
        </html>
        """
        self.text_edit.setHtml(html)

    def generate_decision_analysis(self, crm_data):
        """Generate textual decision analysis based on conditions, line by line, with final professional decision."""
        if not crm_data:
            return "<div class='decision-section'><h2>Decision Analysis</h2><p>No sufficient data for analysis.</p></div>"
        
        analysis_html = "<div class='decision-section'><h2>Decision Analysis</h2><ul>"
        
        # Condition 1: Check if subtracting blank brings into range
        in_range_with_blank = sum(d['lower'] <= (d['sample_val'] - d['blank_val']) <= d['upper'] for d in crm_data)
        total = len(crm_data)
        analysis_html += f"<li>Condition 1: With blank subtraction, {in_range_with_blank}/{total} CRMs are in range. "
        if in_range_with_blank / total >= 0.75:
            analysis_html += "This is sufficient for most cases; prefer blank adjustment over scaling.</li>"
        else:
            analysis_html += "Insufficient; consider additional adjustments.</li>"
        
        # Condition 2: If blank insufficient, try adding dynamic increment to blank
        best_blank_adjust = 0
        best_in_range = 0
        possible_adjusts = [-15, -10, -5, 0, 5, 10, 15]  # Enhanced dynamic: based on magnitude
        for adjust in possible_adjusts:
            in_range = sum(d['lower'] <= (d['sample_val'] - d['blank_val'] + adjust) <= d['upper'] for d in crm_data)
            if in_range > best_in_range:
                best_in_range = in_range
                best_blank_adjust = adjust
        analysis_html += f"<li>Condition 2: Adjusting blank by {best_blank_adjust} (dynamic increment), achieves {best_in_range}/{total} in range. "
        if best_in_range / total >= 0.75:
            analysis_html += "No scaling needed; apply this blank adjustment globally.</li>"
        else:
            analysis_html += "Still insufficient; proceed to scaling.</li>"
        
        # Condition 3: If majority (e.g., 75%) in range already or with blank, no scaling
        initial_in_range = sum(d['lower'] <= d['sample_val'] <= d['upper'] for d in crm_data)
        analysis_html += f"<li>Condition 3: Initially, {initial_in_range}/{total} in range. With blank: {in_range_with_blank}/{total}. "
        if initial_in_range / total >= 0.75 or in_range_with_blank / total >= 0.75:
            analysis_html += "Majority in range; no scaling required. Apply blank subtraction if needed.</li>"
        else:
            analysis_html += "Minority in range; scaling may be necessary.</li>"
        
        # Condition 4: For duplicate CRM IDs, if at least one in range, no scaling
        from collections import defaultdict
        id_groups = defaultdict(list)
        for d in crm_data:
            id_groups[d['id']].append(d)
        good_groups = sum(any(d['lower'] <= d['sample_val'] <= d['upper'] for d in group) for group in id_groups.values())
        analysis_html += f"<li>Condition 4: For duplicate CRM IDs, {good_groups}/{len(id_groups)} groups have at least one in range. "
        if len(id_groups) == 0 or good_groups / len(id_groups) >= 0.75:
            analysis_html += "Sufficient coverage; no scaling needed.</li>"
        else:
            analysis_html += "Insufficient; consider scaling.</li>"
        
        # Condition 5: Different magnitudes
        def get_magnitude(val):
            return np.floor(np.log10(abs(val))) if val != 0 else 0
        mag_groups = defaultdict(list)
        for d in crm_data:
            mag = get_magnitude(d['cert_val'])
            mag_groups[mag].append(d)
        analysis_html += f"<li>Condition 5: {len(mag_groups)} magnitude groups detected. "
        for mag, group in mag_groups.items():
            in_range = sum(d['lower'] <= d['sample_val'] <= d['upper'] for d in group)
            analysis_html += f"Magnitude 10^{mag}: {in_range}/{len(group)} in range. "
        if len(mag_groups) > 1:
            analysis_html += "Apply adjustments selectively to low-performing groups without affecting others.</li>"
        else:
            analysis_html += "Uniform magnitude; global adjustment safe.</li>"
        
        # Condition 6: Multiple wavelengths - select best Soln Conc
        analysis_html += "<li>Condition 6: Assuming single wavelength. If multiple, select the one with Soln Conc in range or closest for scaling.</li>"
        
        # Condition 7: If multiple scales, average where most stay in range (allow 1-2 out)
        max_corr = 0.3  # Define max_corr here to match optimization section
        possible_scales = [d['cert_val'] / d['sample_val'] if d['sample_val'] != 0 else 1 for d in crm_data]
        possible_scales = [s for s in possible_scales if 1 - max_corr <= s <= 1 + max_corr]  # Filter within bounds
        if possible_scales:
            avg_scale = np.mean(possible_scales)
            in_range_with_avg = sum(d['lower'] <= d['sample_val'] * avg_scale <= d['upper'] for d in crm_data)
            analysis_html += f"<li>Condition 7: Average scale {avg_scale:.3f} brings {in_range_with_avg}/{total} in range (allowing 1-2 outliers).</li>"
        else:
            analysis_html += "<li>Condition 7: No valid scales within bounds.</li>"
        
        # Condition 8: Best blank selection
        analysis_html += f"<li>Condition 8: Best blank adjustment selected as {best_blank_adjust}, maximizing in-range to {best_in_range}/{total}.</li>"
        
        analysis_html += "</ul>"
        
        # Optimized correction using differential evolution
        def objective(params):
            blank_adjust, scale = params
            in_range_count = 0
            for d in crm_data:
                adjusted_val = (d['sample_val'] - d['blank_val'] - blank_adjust) * scale
                if d['lower'] <= adjusted_val <= d['upper']:
                    in_range_count += 1
            reg = 0.1 * (abs(blank_adjust) + abs(scale - 1))
            return -in_range_count + reg
        
        avg_cert = np.mean([d['cert_val'] for d in crm_data])
        blank_bounds = (-0.15 * avg_cert, 0.15 * avg_cert)
        scale_bounds = (1 - max_corr, 1 + max_corr)
        bounds = [blank_bounds, scale_bounds]
        
        try:
            res = differential_evolution(objective, bounds)
            if res.success:
                blank_adjust, scale = res.x
                in_range_after = 0
                for d in crm_data:
                    adjusted_val = (d['sample_val'] - d['blank_val'] - blank_adjust) * scale
                    if d['lower'] <= adjusted_val <= d['upper']:
                        in_range_after += 1
                
                # Check if blank only is sufficient
                def objective_blank(p):
                    return objective([p[0], 1])
                res_blank = differential_evolution(objective_blank, [blank_bounds])
                if res_blank.success:
                    blank_adjust_blank = res_blank.x[0]
                    in_range_blank = 0
                    for d in crm_data:
                        adjusted_val = (d['sample_val'] - d['blank_val'] - blank_adjust_blank) * 1
                        if d['lower'] <= adjusted_val <= d['upper']:
                            in_range_blank += 1
                    if in_range_blank >= in_range_after or (abs(scale - 1) > 0.01 and in_range_blank / total >= 0.75):
                        blank_adjust = blank_adjust_blank
                        scale = 1.0
                        in_range_after = in_range_blank
                
                analysis_html += f"<p><strong>Optimized Correction:</strong> Blank adjust: {blank_adjust:.3f}, Scale: {scale:.3f}, achieving {in_range_after}/{total} in range.</p>"
                
                # Final decision
                if in_range_after == total:
                    recommended_action = f"Apply optimized correction: subtract {blank_adjust:.3f} and scale by {scale:.3f}; all data will be in range."
                elif in_range_after > initial_in_range:
                    recommended_action = f"Apply optimized correction: subtract {blank_adjust:.3f} and scale by {scale:.3f}; improves to {in_range_after}/{total} in range."
                else:
                    recommended_action = f"No improvement from optimization; recommend applying blank adjustment of {blank_adjust:.3f} without scaling, achieving {in_range_after}/{total} in range."
            else:
                analysis_html += "<p><strong>Optimized Correction:</strong> Optimization failed.</p>"
                recommended_action = "Optimization failed; recommend manual review to determine appropriate blank or scale adjustments."
        except Exception as e:
            analysis_html += f"<p><strong>Optimized Correction:</strong> Error in optimization: {str(e)}.</p>"
            recommended_action = "Error in optimization; recommend manual review to determine appropriate blank or scale adjustments."
        
        analysis_html += f"<p><strong>Final Decision:</strong> {recommended_action} This ensures the majority of data falls within acceptable ranges with minimal adjustments, aligning with the goal of optimal data integrity.</p></div>"
        
        return analysis_html