import xlsxwriter
import tkinter
top = tkinter.Tk()



nameData = ["Avery Bradley","Jae Crowder","John Holland",
              "R.J. Hunter","Jonas Jerebko","Amir Johnson","Jordan Mickey",
              "Kelly Olynyk","Terry Rozier","Marcus Smart","Jared Sullinger","Isaiah Thomas","Evan Turner","James Young",
              "Tyler Zeller","Bojan Bogdanovic","Markel Brown","Wayne Ellington",
              "Rondae Hollis-Jefferson","Jarrett Jack","Sergey Karasev","Sean Kilpatrick","Shane Larkin","Brook Lopez",
              "Chris McCullough","Willie Reed","Thomas Robinson","Henry Sims","Donald Sloan","Thaddeus Young",
              "Arron Afflalo","Lou Amundson","Thanasis Antetokounmpo","Carmelo Anthony","Jose Calderon",
              "Cleanthony Early","Langston Galloway","Jerian Grant","Robin Lopez","Kyle O'Quinn","Kristaps Porzingis",
              "Kevin Seraphin","Lance Thomas","Sasha Vujacic","Derrick Williams","Tony Wroten","Elton Brand",
              "Isaiah Canaan","Robert Covington","Joel Embiid","Jerami Grant",
              "Richaun Holmes","Carl Landry","Kendall Marshall","T.J. McConnell","Nerlens Noel","Jahlil Okafor","Ish Smith",
              "Nik Stauskas","Hollis Thompson","Christian Wood","Bismack Biyombo","Bruno Caboclo",
              "DeMarre Carroll","DeMar DeRozan","James Johnson","Cory Joseph","Kyle Lowry","Lucas Nogueira",
              "Patrick Patterson","Norman Powell","Terrence Ross","Luis Scola","Jason Thompson","Jonas Valanciunas",
               "Delon Wright","Leandro Barbosa","Harrison Barnes","Andrew Bogut","Ian Clark","Stephen Curry",
              "Festus Ezeli","Draymond Green","Andre Iguodala","Shaun Livingston","Kevon Looney","James Michael McAdoo",
              "Brandon Rush","Marreese Speights","Klay Thompson","Anderson Varejao","Cole Aldrich","Jeff Ayres","Jamal Crawford",
              "Branden Dawson","Jeff Green","Blake Griffin","Wesley Johnson","DeAndre Jordan","Luc Richard Mbah a Moute",
              "Chris Paul","Paul Pierce","Pablo Prigioni","JJ Redick","Austin Rivers","C.J. Wilcox"
            ]

ageData = [25.0,25.0,27.0,22.0,29.0,29.0,21.0,25.0,22.0,22.0,24.0,27.0,27.0,20.0,26.0,27.0,24.0,28.0,21.0,32.0,22.0,26.0
           ,23.0,28.0,21.0,26.0,25.0,26.0,28.0,27.0,30.0,33.0,23.0,32.0,34.0,25.0,24.0,23.0,28.0,26.0,20.0,26.0,28.0
           ,32.0,25.0,23.0,37.0,25.0,25.0,22.0,22.0,22.0,32.0,24.0,24.0,22.0,20.0,27.0,22.0,25.0,20.0,23.0,20.0,29.0
           ,26.0,29.0,24.0,30.0,23.0,27.0,23.0,25.0,36.0,29.0,24.0,24.0,33.0,24.0,31.0,25.0,28.0,26.0,26.0,32.0
           ,30.0,20.0,23.0,30.0,28.0,26.0,33.0,27.0,29.0,36.0,23.0,29.0,27.0,28.0,27.0,29.0,31.0,38.0,39.0
           ,31.0,23.0,25.0
           ]

def cDough():

    workbook = xlsxwriter.Workbook('doughnutChart.xlsx')

    worksheet = workbook.add_worksheet()


    bold = workbook.add_format({'bold': 1})

    headings = ['Name', 'Age']



    worksheet.write_row('A1', headings, bold)


    worksheet.write_column('A2', nameData)
    worksheet.write_column('B2', ageData)


    chart1 = workbook.add_chart({'type': 'doughnut'})


    chart1.add_series({
        'name': 'Doughnut age data',
        'categories': ['Sheet1', 1, 0, 11, 0],
        'values': ['Sheet1', 1, 1, 11, 1],
     })

    chart1.set_title({'name': 'Popular Doughnut Types'})


    chart1.set_style(10)


    worksheet.insert_chart('C2', chart1, {'x_offset': 25, 'y_offset': 10})


    workbook.close()
def cPie():
    workbook = xlsxwriter.Workbook('chart_pie.xlsx')


    worksheet = workbook.add_worksheet()


    bold = workbook.add_format({'bold': 1})


    headings = ['Name', 'Age']




    worksheet.write_row('A1', headings, bold)


    worksheet.write_column('A2', nameData)
    worksheet.write_column('B2', ageData)


    chart1 = workbook.add_chart({'type': 'pie'})


    chart1.add_series({
        'name': 'Pie sales data',
        'categories': ['Sheet1', 9, 0, 21, 0],
        'values': ['Sheet1', 9, 1, 21, 1],
    })

    chart1.set_title({'name': 'Popular Pie Types'})

    chart1.set_style(10)


    worksheet.insert_chart('C2', chart1, {'x_offset': 25, 'y_offset': 10})


    workbook.close()
def cRad():



    workbook = xlsxwriter.Workbook('chart_radar1.xlsx')


    worksheet = workbook.add_worksheet()


    bold = workbook.add_format({'bold': 1})

    headings = ['Number', 'Batch 1', 'Batch 2']

    data = [
        ["Branden Dawson","Jeff Green","Blake Griffin","Wesley Johnson","DeAndre Jordan","Luc Richard Mbah a Moute",
              "Chris Paul","Paul Pierce","Pablo Prigioni","JJ Redick","Austin Rivers","C.J. Wilcox"],
        [23,29,27,28,27,29,31,38,39,31,23,25],
        [20,20,19,18,21,20,19,19,18,20,21,19],
    ]


    worksheet.write_row('A1', headings, bold)

    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])
    worksheet.write_column('C2', data[2])


    chart1 = workbook.add_chart({'type': 'radar'})


    chart1.add_series({
        'name': '=Sheet1!$B$1',
        'categories':'=Sheet1!$A$2:$A$7',
        'values': '=Sheet1!$B$2:$B$7',
    })


    chart1.add_series({
        'name': ['Sheet1', 0, 2],
        'categories': ['Sheet1', 1, 0, 6, 0],
        'values': ['Sheet1', 1, 2, 6, 2],
    })

    chart1.set_title({'name': 'Results of data analysis'})

    chart1.set_x_axis({'name': 'AGES'})

    chart1.set_y_axis({'name': 'Data length (mm)'})

    chart1.set_style(11)


    worksheet.insert_chart('E2', chart1)


    workbook.close()
def cAr():



    workbook = xlsxwriter.Workbook('chart_area.xlsx')


    worksheet = workbook.add_worksheet()



    bold = workbook.add_format({'bold': 1})

    headings = ['Name', 'Age', 'First Prof Age']

    data = [
        ["Chris McCullough","Willie Reed","Thomas Robinson","Henry Sims","Donald Sloan","Thaddeus Young",
              "Arron Afflalo","Lou Amundson","Thanasis Antetokounmpo","Carmelo Anthony","Jose Calderon",],
        [21,26,25,26,28,27,30,33,23,32,34],
        [18,21,21,20,19,19,18,18,19,21,20],
    ]


    worksheet.write_row('A1', headings, bold)

    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])
    worksheet.write_column('C2', data[2])


    chart1 = workbook.add_chart({'type': 'area'})


    chart1.add_series({
        'name': '=Sheet1!$B$1',
        'categories': '=Sheet1!$A$2:$A$12',
        'values': '=Sheet1!$B$2:$B$12',
    })


    chart1.add_series({
        'name': ['Sheet1', 0, 2],
        'categories': ['Sheet1', 1, 0, 12, 0],
        'values': ['Sheet1', 1, 2, 12, 2],
    })

    chart1.set_title({'name': 'Results of data analysis'})

    chart1.set_x_axis({'name': 'Age'})

    chart1.set_y_axis({'name': 'Data length (mm)'})

    chart1.set_style(11)


    worksheet.insert_chart('E2', chart1)


    workbook.close()
def cBar():



    workbook = xlsxwriter.Workbook('chart_bar.xlsx')


    worksheet = workbook.add_worksheet()


    bold = workbook.add_format({'bold': 1})

    headings = ['Name', 'Actual Age', 'First Prof Age']

    data = [
        ["Markel Brown", "Wayne Ellington",
        "Rondae Hollis-Jefferson", "Jarrett Jack", "Sergey Karasev", "Sean Kilpatrick", "Shane Larkin", "Brook Lopez"],
        [24,28,21,32,22,26,23,28],
        [21,24, 18, 20, 18,20,18,19],
    ]


    worksheet.write_row('A1', headings, bold)


    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])
    worksheet.write_column('C2', data[2])


    chart1 = workbook.add_chart({'type': 'bar'})

    
    chart1.add_series({
        'name': '=Sheet1!$B$1',
        'categories':'=Sheet1!$A$2:$A$9',
        'values': '=Sheet1!$B$2:$B$9',
    })


    chart1.add_series({
        'name': ['Sheet1', 0, 2],
        'categories': ['Sheet1', 1, 0, 9, 0],
        'values': ['Sheet1', 1, 2, 9, 2],
    })

    chart1.set_title({'name': 'Results of data analysis'})

    chart1.set_x_axis({'name': 'Age'})

    chart1.set_y_axis({'name': 'Data length (mm)'})

    chart1.set_style(11)


    worksheet.insert_chart('E2', chart1)


    workbook.close()


B = tkinter.Button(top,bg='blue',text ="Doughnut Chart", command = cDough)
B.config(height=5, width=20)
B.grid(row=0,column=1)

B1 = tkinter.Button(top,bg='red',text="Pie Chart",command=cPie)
B1.config(height=5, width=20)
B1.grid(row=1,column=1)
B2 = tkinter.Button(top,bg='purple',text="Radar Chart",command=cRad)
B2.config(height=5, width=20)
B2.grid(row=2,column=1)
B3= tkinter.Button(top,bg='white',text="Area Chart",command=cAr)
B3.config(height=5, width=20)
B3.grid(row=3,column=1)
B4= tkinter.Button(top,bg='green',text="Bar Chart",command=cBar)
B4.config(height=5, width=20)
B4.grid(row=4,column=1)

top.mainloop()