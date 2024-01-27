#importing API to use Google Sheets
import gspread 
from google.oauth2.service_account import Credentials

#connecting to the API, allowing access to Google Drive and Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

credentials = Credentials.from_service_account_file("challenge-credentials.json", scopes=scopes)
client = gspread.authorize(credentials) #authorize using the provided credentials

#open the spreadsheet
sheet = client.open("Engenharia de Software - Desafio Yasmin Soraya").sheet1
total_classes = 60

#loop through rows 4 to 27
for index in range (3, 27): 
    #get values from cells
    absences = int(sheet.cell(index + 1, 3).value) 
    p1 = int(sheet.cell(index + 1, 4).value)
    p2 = int(sheet.cell(index + 1, 5).value)
    p3 = int(sheet.cell(index + 1, 6).value)
    m = (p1 + p2 + p3) / 3

    #calculate student's situation
    def calculate_student_situation(m, absences):
        if absences > 0.25 * total_classes:
            #sould come first as absences result in failure regardless of the average(m)
            return "Reprovado por Falta" 
        elif m < 50:
            return "Reprovado por Nota"
        elif 50 <= m < 70:
            return "Exame Final"
        elif m >= 70:
            return "Aprovado"

    #set the situation value
    situation = calculate_student_situation(m, absences)
    sheet.update_cell(index + 1, 7, situation)

    #if the situation is "Exame Final"
    if situation == "Exame Final":
        naf = 100 - m
        final_approval_grade = (m + naf)/2
        sheet.update_cell(index + 1, 8, final_approval_grade) #fill column 8 with the final approval grade
    else:
        sheet.update_cell(index + 1, 8, 0) #fill column 8 with the number 0