import tkinter as tk
from turtle import right
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import openpyxl as op

# import re
# p = re.compile('[a-z]+')
# print(p.findall("")[0])


class Student():
    def __init__(self, name, class_name, school_term, grade, school_name):
        self.name = name
        self.class_name = class_name
        self.school_term = school_term
        self.grade = grade
        self.school_name = school_name



def exportar_para_pdf():
    pdf_file = "conteudo_exportado.pdf"

    # Criar um novo PDF
    with PdfPages(pdf_file) as pdf:
        # Renderizar e salvar o gráfico
        fig, ax = plt.subplots()
        ax.plot([1, 2, 3, 4], [1, 4, 2, 3])
        ax.set_xlabel('X')
        ax.set_ylabel('Y')
        ax.set_title('Gráfico de Exemplo')
        pdf.savefig()  # Salvar o gráfico no PDF

        # Fechar o gráfico após salvar
        plt.close()

    print("PDF exportado com sucesso.")

# # Criar a janela Tkinter
# root = tk.Tk()
# root.title("Exportar para PDF")

# # Criar um botão para acionar a exportação para PDF
# botao_exportar = tk.Button(root, text="Exportar Gráfico para PDF", command=exportar_para_pdf)
# botao_exportar.pack()

# root.mainloop()

workbook = op.load_workbook('modelo_de_dados_copy.xlsx')
ws = workbook.active

# print(workbook.active.cell(row=1, column=1).value)

# for row in ws.iter_rows(min_row=2, max_col=5):
#     # print(row)
#     for cell in row:
#         print(cell(1, 1))
#     # for i in range(13, 200):
#     # print()

subject = ""
competence = ""

questions_rigth = 0

# student = Student()
total_questions = {
    'lp':0, 
    'mat':0,
    'geo':0,
    'his':0,
    'bio':0,
    'ing':0
}

total_correct_ans = {
    'lp':0, 
    'mat':0,
    'geo':0,
    'his':0,
    'bio':0,
    'ing':0
}

name = ""
class_name = ""
school_term = ""
grade = ""
school = ""

print(ws.max_row)
print(ws.max_column)

for irow in range(2, ws.max_row+1):
    name = ws.cell(row=irow, column=1).value
    class_name = ws.cell(row=irow, column=2).value
    school_term = ws.cell(row=irow, column=3).value
    grade = ws.cell(row=irow, column=4).value
    school = ws.cell(row=irow, column=5).value

    for icolumn in range(6, ws.max_column+1, 9):
        chose_alternative = ws.cell(row=irow, column=icolumn+7).value
        correct_alternative = ws.cell(row=irow, column=icolumn+8).value
        
        subject = ws.cell(row=irow, column=icolumn).value

        match subject:
            case "LP":
                if (correct_alternative == chose_alternative):
                    total_correct_ans["lp"] += 1
                total_questions["lp"] += 1
            case "MAT":
                if (correct_alternative == chose_alternative):
                    total_correct_ans["mat"] += 1
                total_questions["mat"] += 1
            case "GEO":
                if (correct_alternative == chose_alternative):
                    total_correct_ans["geo"] += 1
                total_questions["geo"] += 1
            case "HIS":
                if (correct_alternative == chose_alternative):
                    total_correct_ans["his"] += 1
                total_questions["his"] += 1
            case "BIO":
                if (correct_alternative == chose_alternative):
                    total_correct_ans["bio"] += 1
                total_questions["bio"] += 1
            case "ING":
                if (correct_alternative == chose_alternative) :
                    total_correct_ans["ing"] += 1
                total_questions["ing"] += 1

        if (correct_alternative == chose_alternative):
            questions_rigth += 1
    
    
    print("Correct questions:", questions_rigth)
    
    questions_rigth = 0

    print("Total questions:")
    for key, value in total_questions.items():
        print(key, " : ", total_questions[key])
        total_questions[key] = 0
    
    print("Correct ans:")
    for key, value in total_correct_ans.items():
        print(key, " : ", total_correct_ans[key])
        total_correct_ans[key] = 0

