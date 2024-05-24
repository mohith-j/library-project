# Python Libraries
import pandas as pd
import fitz

filename = 'Empty Labels.pdf' # Template File Path
doc = fitz.open(filename) # For printing values
temp_doc = fitz.open(filename) # For Template
pno=0
count=0
coordn=[(35,110),(35, 140), (335,110),(335,140)] # Vertical Distance between labels = 115 units

excel_file_path = 'Noble Library_Periodical Labels.xlsx' # Path for excel file input
df = pd.read_excel(excel_file_path) # Getting Data from Excel Sheet

for i in range(0,len(df)):
  page = doc[pno]
  if(count<6):
    point1 = fitz.Point((coordn[0][0],coordn[0][1]+(count*117)))
    point2 = fitz.Point((coordn[1][0],coordn[1][1]+(count*117)))
    page.insert_text(point1, df.iloc[i,0], fontname = "Verdana", fontsize = 16, fontfile='Verdana.ttf', color=(1, 1, 1), rotate = 0, )
    if len(df.iloc[i,1]) < 35:
      page.insert_text(point2, df.iloc[i,1], fontname = "Verdana_Bold", fontsize = 12, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
    else:
      sentence1=df.iloc[i,1][:35]+"-"
      sentence2="-"+df.iloc[i,1][35:]
      page.insert_text(point2, sentence1, fontname = "Verdana_Bold", fontsize = 11, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
      page.insert_text(fitz.Point((coordn[1][0],coordn[1][1]+(count*117)+10)), sentence2, fontname = "Verdana_Bold", fontsize = 11, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
  else:
    point1 = fitz.Point((coordn[2][0],coordn[2][1]+((count-6)*117)))
    point2 = fitz.Point((coordn[3][0],coordn[3][1]+((count-6)*117)))
    page.insert_text(point1, df.iloc[i,0], fontname = "Verdana", fontsize = 16, fontfile='Verdana.ttf', color=(1, 1, 1), rotate = 0, )
    if len(df.iloc[i,1]) < 35:
      page.insert_text(point2, df.iloc[i,1], fontname = "Verdana_Bold", fontsize = 12, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
    else:
      sentence1=df.iloc[i,1][:35]+"-"
      sentence2="-"+df.iloc[i,1][35:]
      page.insert_text(point2, sentence1, fontname = "Verdana_Bold", fontsize = 11, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
      page.insert_text(fitz.Point((coordn[3][0],coordn[3][1]+((count-6)*117)+10)), sentence2, fontname = "Verdana_Bold", fontsize = 11, fontfile='Verdana_Bold.ttf', color=(0.549, 0.114, 0.251), rotate = 0, )
  count+=1
  if(count==12):
    new_page = doc.new_page(width=temp_doc[0].rect.width, height=temp_doc[0].rect.height)
    new_page.show_pdf_page(new_page.rect, temp_doc, 0)
    count=0
    pno+=1
      
doc.save("output.pdf")

doc.close()
temp_doc.close()