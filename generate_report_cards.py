import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os

# File paths
EXCEL_FILE = "D:\\Invest4Edu\\student_scores.xlsx"
OUTPUT_DIR = "D:\\Invest4Edu\\ReportCards"

# Function to validate the Excel file structure
def validate_excel_columns(df, required_columns):
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The Excel file must contain the following columns: {missing_columns}")

# Function to generate a report card for a student
def generate_report_card(student_name, student_data, total_score, average_score, output_dir):
    file_name = os.path.join(output_dir, f"report_card_{student_name.replace(' ', '_')}.pdf")
    doc = SimpleDocTemplate(file_name, pagesize=letter)

    # Styles
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    normal_style = styles['BodyText']

    # Title and summary
    elements = [
        Paragraph(f"Report Card: {student_name}", title_style),
        Paragraph(f"<b>Total Score:</b> {total_score}", normal_style),
        Paragraph(f"<b>Average Score:</b> {average_score:.2f}", normal_style),
    ]

    # Table data
    table_data = [["Subject", "Score"]]
    for _, row in student_data.iterrows():
        table_data.append([row['Subject'], row['Score']])

    # Table styling
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)

    # Build PDF
    doc.build(elements)

# Main script
def main():
    try:
        # Read Excel file
        df = pd.read_excel(r"D:\Invest4Edu\student_scores (1).xlsx")

        # Validate columns
        required_columns = ['Name', 'Subject', 'Score']
        validate_excel_columns(df, required_columns)

        # Group data by student
        grouped = df.groupby('Name')

        # Create output directory if it doesn't exist
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)

        # Generate report cards
        for student_name, student_data in grouped:
            total_score = student_data['Score'].sum()
            average_score = student_data['Score'].mean()
            generate_report_card(student_name, student_data, total_score, average_score, OUTPUT_DIR)

        print("Report cards have been generated successfully!")

    except ValueError as ve:
        print(f"Error: {ve}")
    except FileNotFoundError:
        print(f"Error: The file {EXCEL_FILE} does not exist.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
