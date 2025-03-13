import pandas as pd
from reportlab.pdfgen import canvas
import random

file_name=r'new\payments.xlsx'
df=pd.read_excel(file_name)

def simulate_month():
    base_hours = 168
    worked_hours = base_hours
    absent_days = 0
    for _ in range(21):
        if random.random() < 0.02:
            absent_days += 1
    worked_hours -= absent_days * 8
    overtime_hours = 0
    overtime_shifts = random.randint(0, 4)
    for _ in range(overtime_shifts):
        overtime_hours += random.randrange(2, 19, 2)
    worked_hours += overtime_hours
    for i in range(len(df)):
        df.loc[i,'Worked Hours'] = worked_hours
        h_payment=df.iloc[i]['Hourly Payment']
        h_worked=df.iloc[i]['Worked Hours']
        overtime_payment=1.75*h_payment
        if h_worked < 168:
            payment=h_payment*h_worked
            final_payment=int(payment)
            df.loc[i,'Gross Payment'] = final_payment
            df.loc[i,'Net Payment'] = df.loc[i,'Gross Payment'] *0.65
        elif h_worked == 168:
            df.loc[i,'Gross Payment']=df.loc[i,'Gross Payment']
            df.loc[i,'Net Payment'] = df.loc[i,'Gross Payment'] *0.65
        else:
            h_overtime=h_worked-168
            payment=(h_payment*h_worked)+(overtime_payment*h_overtime)
            final_payment=int(payment)
            df.loc[i,'Gross Payment'] = final_payment
            df.loc[i,'Net Payment'] = df.loc[i,'Gross Payment'] *0.65
    print('Month succesfully simulated')

with pd.ExcelWriter(file_name, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, index=False)

def create_pdf(filename, data):
    c = canvas.Canvas(filename)
    c.drawString(100, 750, "Payroll")
    image_path=r'new\images.png'
    try:
        c.drawImage(image_path, 400, 700, width=100, height=100)
    except Exception as e:
        print(f"Error loading image: {e}")
    y = 650 
    for line in data:
        c.drawString(100, y, line)
        y -= 20
    c.save()

def generate_pdf():
    count=0
    for i in range(len(df)):
        count+=1
        name=df.iloc[i]['Name']
        net_payment = df.iloc[i]['Net Payment']
        base_payment = df.iloc[i]['Base Payment']
        gross_payment = df.iloc[i]['Gross Payment']
        health_insurance = df.iloc[i]['Health Insurance']
        retirement_fund = df.iloc[i]['Retirement Fund']
        worked_h=df.iloc[i]['Worked Hours']
        overtime_h= df.iloc[i]['Worked Hours'] - 168
        data=[
            '                                                                                             Mehigh Executive INC',
            '',
            '                                          Your payment receipt',
            '',
            '',
            f'Name: {name}',
            '',
            'IBAN: IBAN0001EXAMPLE',
            '',
            f'Total Payment: {net_payment}',
            '',
            f'Base Payment: {base_payment}',
            '',
            f'Overtime Payment: {net_payment-base_payment}',
            '',
            f'Gross Payment: {gross_payment}',
            '',
            f'Health Insurance: {health_insurance}',
            '',
            f'Retirement Fund: {retirement_fund}',
            '',
            f'Worked Hours: {worked_h}',
            '',
            f'Overtime Hours/Missed Hours (if with -): {overtime_h}',
            '',
            '',
            '',
            '                                                                                          See you next mouth, take care !',
            '',
            '',
            '',
            'For any questions just send an email at mihaitg19@gmail.com for any information',
            '',
        ]
        create_pdf(f"new/payments/Payment {name}.pdf", data)
    print(f'You succesfully created {count} PDFs')

simulate_month()
generate_pdf()