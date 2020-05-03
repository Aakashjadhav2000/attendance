import xlrd
import xlsxwriter

date = input("Enter the date of class:-")
month = int(input("enter the month:-"))

row = 0
column = 0
subject = 8

outworkbook_1 = xlsxwriter.Workbook('blast-{}-{}-2020.xlsx'.format(date,month))
outsheet_1 = outworkbook_1.add_worksheet()

success = outworkbook_1.add_format(
    {
        "bg_color" : "#AAA"
    }
)

while(subject != 0):
    print("\n0.exit \n1.COA\n2.AOA\n3.OS\n4.CG\n5.M4 \n select subject :-")
    subject = int(input())
    if subject == 1:
        get_code = input("Enter the professor code:-")
        if get_code == "coaprop":
            outsheet_1.write(0,column,"AOA",success)
            num = int(input("Enter the number of students:-"))
            for row in range(1,num+1):
                name = input("enter the name:-")
                outsheet_1.write(row , column, name)

    elif subject == 2 :
        get_code = input("enter the professor code:-")
        if get_code == "aoaprop":
            outsheet_1.write(0,column,"COA",success)
            num = int(input("Enter the number of students:-"))
            for row in range(1, num+1):
                name = input("enter the name:-")
                outsheet_1.write(row, column, name)

    elif subject == 3 :
        get_code = input("enter the professor code:-")
        if get_code == "osprop":
            outsheet_1.write(0,column,"OS",success)
            num = int(input("Enter the number of students:-"))
            for row in range(1, num+1):
                name = input("enter the name:-")
                outsheet_1.write(row, column, name)

    elif subject == 4 :
        get_code = input("enter the professor code:-")
        if get_code == "cgprop":
            outsheet_1.write(0,column,"CG",success)
            num = int(input("Enter the number of students:-"))
            for row in range(1, num+1):
                name = input("enter the name:-")
                outsheet_1.write(row, column, name)

    elif subject == 5 :
        get_code = input("enter the professor code:-")
        if get_code == "m4prop":
            outsheet_1.write(0,column,"M4",success)
            num = int(input("Enter the number of students:-"))
            for row in range(1, num+1):
                name = input("enter the name-")
                outsheet_1.write(row, column, name)

    elif subject == 0 :
        break

    column = column + 1

outworkbook_1.close()
