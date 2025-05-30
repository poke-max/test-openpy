from flask import Flask, render_template, request, flash, redirect, url_for, send_file, make_response
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import re

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your_secret_key_here')  # Use environment variable for secret key

# Use absolute path for deployment
EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'layout', 'PLANILLA.xlsx')

def excel_col_to_index(col_str):
    """Convert Excel column letter to 0-based index"""
    col_str = col_str.upper()
    result = 0
    for char in col_str:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Get form data
            sheet_name = request.form['sheet_name']
            cell = request.form['cell']
            new_value = request.form['new_value']

            if not os.path.exists(EXCEL_FILE):
                flash('El archivo Excel no se encuentra en la ubicación especificada', 'error')
                return redirect(url_for('index'))

            # Read the Excel file
            excel_file = pd.ExcelFile(EXCEL_FILE)
            
            # Check if sheet exists
            if sheet_name not in excel_file.sheet_names:
                flash('¡Hoja no encontrada!', 'error')
                return redirect(url_for('index'))

            # Read the specific sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Parse cell reference (e.g., 'A1' to row and column)
            match = re.match(r'([A-Za-z]+)(\d+)', cell)
            if not match:
                flash('Formato de celda inválido. Use formato como "A1", "B2", etc.', 'error')
                return redirect(url_for('index'))
                
            col_letter = match.group(1)
            row_num = int(match.group(2))
            
            # Convert to 0-based indices
            col_idx = excel_col_to_index(col_letter)
            row_idx = row_num - 1
            
            # Validate indices
            if row_idx < 0 or row_idx >= len(df) or col_idx < 0 or col_idx >= len(df.columns):
                flash('La celda especificada está fuera de los límites de la hoja', 'error')
                return redirect(url_for('index'))
            
            # Update the cell value
            df.iloc[row_idx, col_idx] = new_value
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'PLANILLA_actualizada_{timestamp}.xlsx'
            
            # Save to memory buffer
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)
            
            # Create response with cache headers
            response = make_response(send_file(
                output,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ))
            
            # Set cache headers
            response.headers['Cache-Control'] = 'public, max-age=31536000'  # Cache for 1 year
            response.headers['Expires'] = '31536000'
            
            return response
            
        except Exception as e:
            flash(f'Ha ocurrido un error: {str(e)}', 'error')
        
        return redirect(url_for('index'))
    
    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port) 