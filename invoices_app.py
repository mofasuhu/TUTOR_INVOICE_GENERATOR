import pandas as pd
import re
from bidi.algorithm import get_display
import arabic_reshaper
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from warnings import filterwarnings

filterwarnings("ignore", category=DeprecationWarning)

def normalize_space(string):
    return str(re.sub(r'\s+', ' ', string)).strip(' ').strip('\n').replace(" - - - - - - - "," - ").replace(" - - - - - - "," - ").replace(" - - - - - "," - ").replace(" - - - - "," - ").replace(" - - - "," - ").replace(" - - "," - ").strip(' - ')
    
def pdf_ar_fix(text):
    return get_display(arabic_reshaper.reshape(text))

# Define a function to check if a grade is in Arabic
def is_arabic_grade(grade):
    return any(char.isalnum() and not char.isdigit() for char in grade) and grade not in ["Compensation", "CRM", "Content Task"]

# Define a function to consolidate data for each tutor
def consolidate_tutor_data(tutor_df):
    # Create a dictionary to hold the consolidated data
    try:
        CST_price = tutor_df['Content Special Tasks Price'][tutor_df['Content Special Tasks Price']!=0].tolist()
        CST_total = tutor_df['Content Special Total'][tutor_df['Content Special Total']!=0].tolist()
        CRM_dur = tutor_df['CRM Duration'][tutor_df['CRM Duration']!=0].tolist()
        CRM_price = tutor_df['CRM Price'][tutor_df['CRM Price']!=0].tolist()
        CRM_pay = tutor_df['CRM Payment'][tutor_df['CRM Payment']!=0].tolist()
        Comp_total = tutor_df['Total Compensation'][tutor_df['Total Compensation']!=0].tolist()
        Demo_price = tutor_df['Demo Price'][tutor_df['Demo Price']!=0].tolist()
        Demo_total = tutor_df['Demo Total'][tutor_df['Demo Total']!=0].tolist()
        Demo_month = tutor_df['Demo Month'][tutor_df['Demo Month']!=""].tolist()       
        data = {
            'ID': tutor_df['ID'].iloc[0],
            'Full Name': tutor_df['Full Name'].iloc[0],
            'Address': tutor_df['Address'].iloc[0],
            'Invoice Number': tutor_df['Invoice Number'].iloc[0],
            'Invoice Date': tutor_df['Invoice Date'].iloc[0],
            'Mobile': tutor_df['Mobile'].iloc[0],
            'Email Address': tutor_df['Email Address'].iloc[0],
            'Subjects': normalize_space(" - ".join(list(set(tutor_df['Subject 1'].unique().tolist()+tutor_df['Subject 2'].unique().tolist()+tutor_df['Subject 3'].unique().tolist())))),
            'Accrual Month': tutor_df['Accrual Month'].iloc[0],        
            'Total Sessions': tutor_df['Total Sessions per Grade'].sum(),
            'Content Special Tasks': tutor_df['Content Special Tasks'].sum(),
            'Content Special Tasks Price': CST_price[0] if CST_price else 0,
            'Content Special Tasks Total': CST_total[0] if CST_total else 0,        
            'CRM Duration': round(CRM_dur[0],2) if CRM_dur else 0,
            'CRM Price': CRM_price[0] if CRM_price else 0,    
            'CRM Payment': round(CRM_pay[0]) if CRM_pay else 0,
            'Demo No.': tutor_df['Demo No.'].sum(),
            'Demo Price':  Demo_price[0] if Demo_price else 0,
            'Demo Total':  Demo_total[0] if Demo_total else 0,
            'Demo Month': Demo_month[0] if Demo_month else "",
            'Total Compensation': Comp_total[0] if Comp_total else 0,
            'Total Salary': tutor_df['Total Salary'].sum(),
            'Bank Account Name': tutor_df['Bank Account Name'].iloc[0],
            'Bank Name': tutor_df['Bank Name'].iloc[0],
            'Bank Address': tutor_df['English Bank Address'].iloc[0],
            'Account Number': tutor_df['Account Number'].iloc[0],
            'Bank Account Number (IBAN)': tutor_df['Bank Account Number (IBAN)'].iloc[0],
            'Swift': tutor_df['Swift'].iloc[0]
        }
        
        # Create columns for each unique "Grade - Subject" pair and sum the corresponding values
        pairs = {}
        for _, row in tutor_df.iterrows():
            for i in range(1, 4):
                subject = row[f'Subject {i}']
                grade = row['Grade']
                if subject and is_arabic_grade(grade):
                    pair = f"{subject} {grade}"
                    if pair not in pairs:
                        pairs[pair] = {
                            'Total Sessions': 0,
                            'Session Price': 0,
                            'Total Sessions Price': 0
                        }
                    pairs[pair]['Total Sessions'] += row[f'Total Sessions {i}']
                    pairs[pair]['Session Price'] += row[f'Session Price {i}']
                    pairs[pair]['Total Sessions Price'] += row[f'Total Sessions Price {i}']

        for idx, (pair, values) in enumerate(pairs.items()):
            data[str(idx + 1)] = pair
            data[f'Total Sessions {str(idx + 1)}'] = values['Total Sessions']
            data[f'Session Price {str(idx + 1)}'] = values['Session Price']
            data[f'Total Sessions Price {str(idx + 1)}'] = values['Total Sessions Price']

        return data
    
    except Exception as e:
        print(f"Error consolidating tutor data: {e} in {tutor_df}.")
        return {}

# Function to calculate column widths
def calculate_column_widths(data, font_name, font_size):
    try:
        widths = []
        for i in range(len(data[0])):
            col_width = max([pdfmetrics.stringWidth(str(row[i]), font_name, font_size) for row in data])
            widths.append(col_width + 10)  # Add padding
        total_width = sum(widths)
        return widths, total_width
    except Exception as e:
        print(f"Error calculating column widths: {e}")
        return [], 50

try:
    file_path = 'tutorlist.xlsx'
    df = pd.read_excel(file_path).fillna("")
    df = df.rename(columns=lambda x: normalize_space(x))
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date']).dt.strftime('%d %B %Y')
    df[['Session Price 1','Session Price 2','Session Price 3']] = df[['Session Price 1','Session Price 2','Session Price 3']].astype(int)    

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(normalize_space)
            
    # Group the data by tutor ID
    grouped = df.groupby('ID')

    # Consolidate data for each tutor
    consolidated_data = [consolidate_tutor_data(tutor_df) for _, tutor_df in grouped]

    # Create a new DataFrame with the consolidated data
    consolidated_df = pd.DataFrame(consolidated_data).fillna("--")

    consolidated_df[['Total Salary', 'Session Price 1','Session Price 2','Session Price 3','Session Price 4',
                     'Session Price 5','Session Price 6','Session Price 7',
                     'Total Sessions Price 1','Total Sessions Price 2','Total Sessions Price 3','Total Sessions Price 4',
                     'Total Sessions Price 5','Total Sessions Price 6','Total Sessions Price 7', 
                     'Total Sessions 1','Total Sessions 2','Total Sessions 3','Total Sessions 4',
                     'Total Sessions 5','Total Sessions 6','Total Sessions 7']] = consolidated_df[['Total Salary', 'Session Price 1','Session Price 2','Session Price 3','Session Price 4',
                     'Session Price 5','Session Price 6','Session Price 7',
                     'Total Sessions Price 1','Total Sessions Price 2','Total Sessions Price 3','Total Sessions Price 4',
                     'Total Sessions Price 5','Total Sessions Price 6','Total Sessions Price 7', 
                     'Total Sessions 1','Total Sessions 2','Total Sessions 3','Total Sessions 4',
                     'Total Sessions 5','Total Sessions 6','Total Sessions 7']].astype(str).replace('--','0').astype(float).astype(int)
    
    # Write the consolidated data to a new Excel sheet
    output_file_path = 'consolidated_tutor_data.xlsx'
    consolidated_df.to_excel(output_file_path, index=False)

    print(f"Consolidated data has been written to {output_file_path}")

except FileNotFoundError as fnf_error:
    print(f"File not found: {fnf_error}")
    consolidated_df = pd.DataFrame([])
except pd.errors.EmptyDataError as ede_error:
    print(f"Empty data error: {ede_error}")
    consolidated_df = pd.DataFrame([])
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    consolidated_df = pd.DataFrame([])

pdfmetrics.registerFont(TTFont('NotoSerif', 'fonts/NotoSerif-Bold.ttf'))
pdfmetrics.registerFont(TTFont('MyNoto', 'fonts/NotoNaskhArabic-Regular.ttf'))
pdfmetrics.registerFont(TTFont('MyNotoBold', 'fonts/NotoNaskhArabic-Bold.ttf'))
    
digit_columns = [col for col in consolidated_df.columns if col.isdigit()]

counter = 0
for idx, row in consolidated_df.iterrows():
    counter += 1
    print(f'\r{counter}', end=' ')
    pdf_file_path = f"PDFs/{normalize_space(row['Full Name'])}_{row['ID']}.pdf"
    c = canvas.Canvas(pdf_file_path, pagesize=(A4[0] + 20, A4[1] + 20))
    c.translate(10, 10)
    width, height = A4
    # header
    c.setFont('MyNotoBold', 12)
    c.drawString(25, height - 25, 'Invoice')

    # Box
    c.setLineJoin(1)  # round
    c.setLineWidth(1.5)
    box_x, box_y, box_width, box_height = 25, height - 70, (width / 2 - 30), 20

    try:
        y_factor = 0
        c.setFont('MyNotoBold', 10)
        c.drawString(30, height - 45 - y_factor, 'Name:')
        c.rect(box_x, box_y - y_factor, box_width, box_height, stroke=1, fill=0)
        c.setFont('MyNoto', 9)
        c.drawString(box_x + 5, box_y - y_factor + 6.25, pdf_ar_fix(row['Full Name']))

        y_factor = 45 * 1
        c.rect(box_x, box_y - y_factor - 10, box_width, box_height + 10, stroke=1, fill=0)
        styles = ParagraphStyle(name='', fontName='MyNoto', fontSize=9, textColor='black')
        text = pdf_ar_fix(row['Address'])
        paragraph = Paragraph(text, styles, encoding='utf-8')
        frame = Frame(box_x, box_y - y_factor - 10 - 5, box_width - 5, box_height + 18, showBoundary=0)
        frame.addFromList([paragraph], c)
        c.setFont('MyNotoBold', 10)
        c.drawString(30, height - 45 - y_factor, 'Address:')

        y_factor = 10 + 45 * 2
        c.setFont('MyNotoBold', 10)
        c.drawString(30, height - 45 - y_factor, 'Mobile Number:')
        c.rect(box_x, box_y - y_factor, box_width, box_height, stroke=1, fill=0)
        c.setFont('MyNoto', 9)
        c.drawString(box_x + 5, box_y - y_factor + 6.25, pdf_ar_fix(row['Mobile']))

        y_factor = 10 + 45 * 3
        c.setFont('MyNotoBold', 10)
        c.drawString(30, height - 45 - y_factor, 'Email Address:')
        c.rect(box_x, box_y - y_factor, box_width, box_height, stroke=1, fill=0)
        c.setFont('MyNoto', 9)
        c.drawString(box_x + 5, box_y - y_factor + 6.25, pdf_ar_fix(row['Email Address']))

        y_factor = 10 + 45 * 4
        c.setFont('MyNotoBold', 10)
        c.drawString(30, height - 45 - y_factor, 'To:')
        c.rect(box_x, box_y - y_factor - 30, box_width, box_height + 30, stroke=1, fill=0)
        c.setFont('MyNoto', 9)

        text = c.beginText()
        text.setTextOrigin(box_x + 5, box_y - y_factor + 6.25)
        text.setFont('MyNoto', 9)
        text.setLeading(15)
        text.textLines("Nagwa Limited\nYork House, 41 Sheet Street, Windsor, SL4 1DD\nUNITED KINGDOM")
        c.drawText(text)

        y_factor = 10 + 45 * 4
        c.rect(width - box_x, box_y - y_factor - 30, -box_width, box_height + 30, stroke=1, fill=0)
        styles = ParagraphStyle(name='', fontName='MyNoto', fontSize=9, textColor='black', leading=14)
        text = f"{pdf_ar_fix(row['Subjects'])}<br />{pdf_ar_fix(row['Accrual Month'])}"
        paragraph = Paragraph(text, styles, encoding='utf-8')
        frame = Frame(width - box_x - box_width, 104, box_width - 5, 500, showBoundary=0)
        frame.addFromList([paragraph], c)
        c.setFont('MyNotoBold', 10)
        c.drawString((width / 2) + 10, height - 45 - y_factor, 'For:')

        c.setFont('MyNotoBold', 7)
        c.drawRightString(width - box_x - 125, height - 45, 'Unique Invoice Number:')
        c.rect(width - box_x - 5, height - 50, -115, 15, stroke=1, fill=0)
        c.setFont('MyNoto', 9)
        c.drawCentredString(width - box_x - 5 - 115 / 2, height - 45.5, row['Invoice Number'])

        c.setFont('MyNotoBold', 7)
        c.drawRightString(width - box_x - 125, height - 45 - 15, 'Invoice Date:')
        c.rect(width - box_x - 5, height - 50 - 15, -115, 15, stroke=1, fill=0)
        c.setFont('MyNoto', 9)
        c.drawCentredString(width - box_x - 5 - 115 / 2, height - 45.5 - 15, row['Invoice Date'])
        c.setFont('MyNoto', 7)
        c.drawCentredString(width - box_x - 5 - 115 / 2, height - 45.5 - 30, '(Last day of the month being invoiced)')
        
    except Exception as e:
        print(f"Error drawing invoice details: {e}")

    page_width = 50
    table_y = 50
    # Define table data
    table_data = [
        ['Description', 'Amount', 'Rate\n(State currency)', 'Rate\n(State currency)']
    ]
    try:
        # Add dynamic rows for each subject-grade pair
        for idx in digit_columns:
            row_data = [
                row[f'{idx}'],
                str(row[f'Total Sessions {idx}']),
                'EGP ' + str(row[f'Session Price {idx}']),
                'EGP ' + str(row[f'Total Sessions Price {idx}'])
            ]
            table_data.append(row_data)

        # Add constant rows
        constant_rows = [
            ['Compensation', '--', '--', f'EGP {row["Total Compensation"]}'],
            ['Content Special Tasks', f'{row["Content Special Tasks"]}', f'EGP {row["Content Special Tasks Price"]}', f'EGP {row["Content Special Tasks Total"]}'],
            ['CRM', f'{row["CRM Duration"]} hr', 'EGP 200', f'EGP {row["CRM Payment"]}'],
            [f'Demo {row["Demo Month"]}', f'{row["Demo No."]}', f'EGP {row["Demo Price"]}', f'EGP {row["Demo Total"]}'],
            ['Total', '--', '--', f'EGP {row["Total Salary"]}']
        ]
        table_data.extend(constant_rows)

        table_data = [list(map(pdf_ar_fix, table)) for table in table_data]

        # Calculate column widths
        column_widths, total_width = calculate_column_widths(table_data, 'MyNoto', 10)

        # Calculate scaling factor to fit table within page margins
        page_width = width - 50  # Subtracting 25 units from both sides for margins
        scaling_factor = page_width / total_width
        # Adjust column widths based on scaling factor
        adjusted_column_widths = [w * scaling_factor for w in column_widths]
        # Create table
        table = Table(table_data, colWidths=adjusted_column_widths)

        # Apply table style
        # (col 0-index, row 0-index) for starting cell and end cell (indexing is from top left to bottom right) 
        # and -1 is last index like python list indexing 
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'MyNotoBold'),
            ('FONTNAME', (0, 1), (-1, -1), 'MyNoto'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertically align text to the middle
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1.5, colors.black),  # Add black grid lines
        ]))

        # Draw the table below the existing content
        table_y = height - (45 * 5 + 350 + 5)  # Adjust this value based on your layout needs
        table.wrapOn(c, width - 50, table_y)  # Ensure table is wrapped before drawing
        table.drawOn(c, 25, table_y)

    except Exception as e:
        print(f"Error creating table: {e}")

    try:
        c.setFont('MyNotoBold', 10)
        c.drawString(30, 245, 'Payment Details:')

        # Define the additional table data
        additional_table_data = [
            ['Account Name', row["Bank Account Name"]],
            ['Bank Name', row["Bank Name"]],
            ['Bank Address', row["Bank Address"]],
            ['□ Account Number\n□ IBAN\n(Tick which it is)', f'{row["Account Number"]}\n{row["Bank Account Number (IBAN)"]}\n'],
            ['□ Swift\n□ BIC\n□ Routing Number\n□ Sort Code\n(Tick which it is)', f'{row["Swift"]}\n\n\n\n'],
            ['Account Type', 'EGP'],
            ['Payment Method', 'Transfer']
        ]
        additional_table_data = [list(map(pdf_ar_fix, table)) for table in additional_table_data]

        # Create the additional table
        additional_table = Table(additional_table_data, colWidths=[page_width * 0.2, page_width * 0.8])

        # Apply style to the additional table
        additional_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'NotoSerif'),
            ('FONTNAME', (-1, 0), (-1, -1), 'MyNoto'),
            ('BACKGROUND', (0, 0), (0, -1), colors.whitesmoke),
            ('BACKGROUND', (1, 0), (1, -1), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (0, -1), 10),
            ('FONTSIZE', (1, 0), (1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1.5, colors.black),  # Add black grid lines
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertically align text to the middle
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))

        # Draw the additional table below the first table
        additional_table_y = table_y - (len(table_data) * 20 - 25)  # Adjust this value based on your layout needs
        additional_table.wrapOn(c, width - 50, additional_table_y)
        additional_table.drawOn(c, 25, additional_table_y)

    except Exception as e:
        print(f"Error creating additional table: {e}")

    c.showPage()
    c.save()
    print(f"PDF saved: {pdf_file_path}")
