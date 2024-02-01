from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import gspread
from datetime import date
import re
from decouple import config


def nombre_corto(name):
    apellido=name.split(",")[0].strip().split(" ")[0]
    nombre=name.split(",")[1].strip().split(" ")[0]
    return f"{nombre} {apellido}"


def month_num_to_word(num,option):
    if option==1:
        meses=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
    if option==2:
        meses=['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE',"DICIEMBRE"]
    return meses[num-1]


def get_month(word):
    if word.lower()=="today":
        today=date.today()
        month=int(today.strftime("%d/%m/%Y").split("/")[1])
    else:
        month=int(word)
    return month
        

def get_observadores(mes,lugar):
    if lugar=="colegio":
        PATH="C:/Users/gquintero/Desktop/python/chromedriver.exe"
    else:    
        PATH="C:/Users/PC/Desktop/python/chromedriver.exe"   

    driver=webdriver.Chrome(PATH)

    target_month=month_num_to_word(mes,1)
    target_month_num=f"{mes:02}"

    driver.get("https://colegionuevayork.phidias.co/poll/consolidate/people?poll=138&section=227")
    search=driver.find_element(By.ID,value="autofocus")
    search.send_keys(config("phidias_usuario"))
    search=driver.find_element(By.NAME,value="_height")
    search.send_keys(config("phidias_password"))
    search.send_keys(Keys.RETURN)
    estudiantes=driver.find_elements(By.XPATH, value="//tr[@class='m_student_row ']")

    #find the values and create a dic key=name, value=link
    estudiantes_dic={}
    for estudiante in estudiantes:
        link=estudiante.find_element(By.TAG_NAME, value="a")
        name=estudiante.find_element(By.TAG_NAME, value="h2")
        date=estudiante.find_elements(By.TAG_NAME, value="a")[3]
        if date.text.split(" ")[0]==target_month or "hace" in date.text.split(" ")[0]:
            estudiantes_dic[name.text]=link.get_attribute("href")

    #create a dic key=name, value= list of dics with codigo, fecha, responsable
    observadores={}
    for estudiante in estudiantes_dic:
        driver.get(estudiantes_dic[estudiante])
        cuerpo=driver.find_elements(By.TAG_NAME, value="tbody")[2]
        anotaciones=cuerpo.find_elements(By.TAG_NAME, value="tr")
        for anotacion in anotaciones:
            fecha=anotacion.find_elements(By.TAG_NAME, value="td")[2].text.split(" ")[0]
            month=fecha.split("/")[1]
            responsable=anotacion.find_elements(By.TAG_NAME, value="td")[3].text
            if month==target_month_num:
                type=anotacion.find_elements(By.TAG_NAME, value="td")[6].text   #type
                code=""
                if type=="Leve":
                    code=anotacion.find_elements(By.TAG_NAME, value="td")[7].text.split(" ")[0]
                elif type=="Grave":
                    code=anotacion.find_elements(By.TAG_NAME, value="td")[8].text.split(" ")[0]
                else:
                    code=anotacion.find_elements(By.TAG_NAME, value="td")[9].text.split(" ")[0]
                if nombre_corto(estudiante) in observadores:
                    observadores[nombre_corto(estudiante)].append({"codigo":code,"fecha":fecha,"responsable":nombre_corto(responsable)})
                else:
                    observadores[nombre_corto(estudiante)]=[{"codigo":code,"fecha":fecha,"responsable":nombre_corto(responsable)}]
    driver.close()
    return observadores


def editar_drive_tarde(month,curso,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="0"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.1":
                    counter+=1
                    comentario+=f"{estudiante} ({anotacion['responsable']}),  "
            comentario_counter=str(counter)+"  "+comentario 

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet("Puntualidad Bto")
    col=hoja.find(month)
    row=hoja.find(curso)
    hoja.update_cell(row.row,col.col,comentario_counter)
    print("Se actualizo el drive de llegadas tarde")


def editar_drive_celular(month,curso,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="0"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.11":
                    counter+=1
                    comentario+=f"{estudiante},  "
            comentario_counter=str(counter)+"  "+comentario 

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet("Celular Bto")
    col=hoja.find(month)
    row=hoja.find(curso)
    hoja.update_cell(row.row,col.col,comentario_counter)
    print("Se actualizo el drive de uso de celular")


def editar_drive_uniforme(month,curso,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="0"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.33":
                    counter+=1
                    comentario+=f"{estudiante},  "
            comentario_counter=str(counter)+"  "+comentario 

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet(" Uniformes Bto")
    col=hoja.find(month)
    row=hoja.find(curso)
    hoja.update_cell(row.row,col.col,comentario_counter)
    print("Se actualizo el drive de uniforme")


def editar_drive_escuela_virtual(month,curso,lugar):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    google_sheet=sa.open("Escuela virtual 10b 2022")
    hoja=google_sheet.worksheet("Hoja 1")
    col=hoja.find(month)
    row=hoja.find(curso)
    num_escuelas=hoja.cell(25,col.col).value
    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet("Esc. Virtual")
    col=hoja.find(re.compile(r"."+month,re.DOTALL))#this finds the first cell with anything before month
    row=hoja.find(curso)
    hoja.update_cell(row.row,col.col-1,"23")
    hoja.update_cell(row.row,col.col,num_escuelas)
    #porcentage=num_escuelas/23
    hoja.update_cell(row.row,col.col+1,f"{int(num_escuelas)*100//23}%")
    print("Se actualizo el drive de escuelas virtuales")


def editar_Hice_9(month,lugar,curso,num_tareas):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
    
    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet("HICE Noveno")
    titulo_mes=hoja.find(f"NOVENO - {month}")
    curso_posicion=["9A","9B","9C","9D","9E"]
    adition=curso_posicion.index(curso)+2
    hoja.update_cell(titulo_mes.row+adition,titulo_mes.col+1,num_tareas)#tareas colocadas
    print(f"Se actualizo el drive tareas de {curso}")


def editar_Hice_diploma(month,lugar,curso,num_tareas):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
    
    google_sheet=sa.open("2023 - INDICADORES MENSUALES")
    hoja=google_sheet.worksheet("HICE Diploma")
    titulo_mes=hoja.find(f"DIPLOMA - {month}")
    if curso==10:
        curso_posicion=["10A","10B","10C","10D","10E"]
        increase=2
    elif curso==11:
        curso_posicion=["11A","11B","11C"]
        increase=7
    for curso in curso_posicion:
        adition=curso_posicion.index(curso)+increase
        tareas=hoja.cell(titulo_mes.row+adition,22).value 

        if tareas:
            if "Gus" in tareas:   
                pass   
            else:
                tareas+=f" - Gus({num_tareas})"
        else:
            tareas=f"Gus({num_tareas})"
    
        hoja.update_cell(titulo_mes.row+adition,22,tareas)#tareas colocadas
    print(f"Se actualizo el drive de tareas de {curso[:2]}")


def main():
    mes=get_month("11") #"today" para el mes actual, "#" para el mes #.  Asi, "7" para el mes 7
    curso="11B"
    lugar="colegio"
    
    month=month_num_to_word(mes,2)
    observadores=get_observadores(mes,lugar)
    editar_drive_tarde(month,curso,observadores,lugar)
    editar_drive_celular(month,curso,observadores,lugar)
    editar_drive_uniforme(month,curso,observadores,lugar)
    editar_drive_escuela_virtual(month,curso,lugar)
    editar_Hice_9(month,lugar,"9B",0)
    editar_Hice_9(month,lugar,"9C",0)
    editar_Hice_diploma(month,lugar,10,0)
    editar_Hice_diploma(month,lugar,11,0)


if __name__ == "__main__":
    main()