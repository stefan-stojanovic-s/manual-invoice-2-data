import fitz
import csv
import pandas as pd
import re
#pymupdf package - fitz
#working with dataframe and then exporting it as a csv
df=pd.DataFrame(data=None)

#parsing the doc
ifile = "rn.pdf"
doc = fitz.open(ifile)
page_count = doc.pageCount
page = 0
text = ''
parser=[]

#it goes through the pages,we get a list of pages of text
while (page < page_count):
    p = doc.loadPage(page)
    page += 1
    text=text+p.getText()
    parser.append(p.getText().split('\n'))

#varriables that are only on the first page
lice=parser[0][1]
broj_fakture=parser[0][parser[0].index('Faktura')+1]
datum_otpreme=parser[0][parser[0].index('Otpremnica:')+1]

#adding to dataframe
df=df.append({'Poslovno lice':lice,
           'Broj fakture':broj_fakture,
           'Datum otpreme':datum_otpreme},ignore_index=True)


#for for number of pages
for i in range(page_count):
    start=parser[i].index('RBr Šifra / Bar ') + 15
    #when the second page enteres loop
    try:
        end=parser[i].index('Poziv na broj:')
    except ValueError:
        end=len(parser[i])
    #for for going through the parser
    for j in range(start,end):
        #regex for br,sifru ex.:1 0506983
        x=re.search('^(\d){1,4}\s.+$',parser[i][j])

        if x != None:
            #assigning each value to its corresponding columns
            br,sifra=parser[i][j].split(' ')
            naziv=parser[i][j+1]
            jm=parser[i][j+2]
            cena=parser[i][j+3]
            cena_sa_pop=parser[i][j+4]
            kol=parser[i][j+5]
            pdv=kol=parser[i][j+6]
            cena_bez_pdv=parser[i][j+7]
            pdv_posto=parser[i][j+8]
            popust=parser[i][j+9]
            barkod=parser[i][j+10]

            #checking if caught the wrong line with regex :)
            #pattern for ammount of products 
            pattern_ammount=re.search('^(\w){2,3}$',jm)
            if pattern_ammount == None:
                naziv=naziv + parser[i][j+2]
                jm=parser[i][j+3]
                cena=parser[i][j+4]
                cena_sa_pop=parser[i][j+5]
                kol=parser[i][j+6]
                pdv=kol=parser[i][j+7]
                cena_bez_pdv=parser[i][j+8]
                pdv_posto=parser[i][j+9]
                popust=parser[i][j+10]
                barkod=parser[i][j+11]
            
            #adding to dataframe then exporting to csv
            df = df.append({'Rbr':br,
                            'Sifra':sifra,
                            'Barkod':barkod,
                            'Naziv':naziv,
                            'J/M':jm,
                            'Kolicina':kol,
                            'Cena':cena,
                            'Pop %':popust,
                            'Cena-Popust':cena_sa_pop,
                            'PDV':pdv,
                            'PDV %':pdv_posto,
                            'Vrednost bez PDV':cena_bez_pdv},ignore_index=True)




df.to_excel('faktura.xlsx')
