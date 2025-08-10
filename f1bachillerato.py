import pandas as pd # type: ignore
from corregir_nombres import corregir_nombre
import numpy as np # type: ignore

df = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name='pb')
df_g = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx',sheet_name ='g')
dfn = pd.read_excel('C:/Users/Admin/Desktop/GK/BASE DE DATOS 2025.xlsx', sheet_name='GK2025')

df['Estudiante'] = df['Estudiante'].apply(corregir_nombre)
df_g['ESTUDIANTE'] = df_g['ESTUDIANTE'].apply(corregir_nombre)
dfn['ESTUDIANTE'] = dfn['ESTUDIANTE'].apply(corregir_nombre)

filasnot, _ = dfn.shape

# define el tamaño de la base de la planeacion
tamaño = df.shape
filasdf,columnasdf = tamaño

Areas = ['C1','C2','S1','S2','L','M1','M2']

for a in Areas:

    #Lista de modulos
    LMOD1l = ["LMOD1"] + df[df.iloc[:, 2] == a].iloc[:, 0].tolist()
    LMOD2l = ["LMOD2"] + df[df.iloc[:, 7] == a].iloc[:, 0].tolist()
    MMOD1l = ["MMOD1"] + df[df.iloc[:, 3] == a].iloc[:, 0].tolist()
    MMOD2l = ["MMOD2"] + df[df.iloc[:, 8] == a].iloc[:, 0].tolist()
    WMOD1l = ["WMOD1"] + df[df.iloc[:, 4] == a].iloc[:, 0].tolist()
    WMOD2l = ["WMOD2"] + df[df.iloc[:, 9] == a].iloc[:, 0].tolist()
    JMOD1l = ["JMOD1"] + df[df.iloc[:, 5] == a].iloc[:, 0].tolist()
    JMOD2l = ["JMOD2"] + df[df.iloc[:, 10] == a].iloc[:, 0].tolist()
    VMOD1l = ["VMOD1"] + df[df.iloc[:, 6] == a].iloc[:, 0].tolist()
    VMOD2l = ["VMOD2"] + df[df.iloc[:, 11] == a].iloc[:, 0].tolist()
    
    LMOD1 = pd.DataFrame(LMOD1l)
    LMOD2 = pd.DataFrame(LMOD2l)
    MMOD1 = pd.DataFrame(MMOD1l)
    MMOD2 = pd.DataFrame(MMOD2l)
    WMOD1 = pd.DataFrame(WMOD1l)
    WMOD2 = pd.DataFrame(WMOD2l)
    JMOD1 = pd.DataFrame(JMOD1l)
    JMOD2 = pd.DataFrame(JMOD2l)
    VMOD1 = pd.DataFrame(VMOD1l)
    VMOD2 = pd.DataFrame(VMOD2l)
    
    df_nombres = pd.concat([LMOD1, LMOD2, MMOD1, MMOD2, WMOD1, WMOD2, JMOD1, JMOD2, VMOD1, VMOD2], ignore_index=True)

    # Crear un diccionario clave ESTUDIANTE y valor GRUPO
    grupo_map = dict(zip(df_g['ESTUDIANTE'], df_g['GRUPO']))
    grado_map = dict(zip(df_g['ESTUDIANTE'], df_g['GRADO']))

    # Asignar grupo y grado correspondiente a cada estudiante de df_nombres
    df_nombres['1'] = df_nombres.iloc[:, 0].map(grupo_map)
    df_nombres['2'] = df_nombres.iloc[:, 0].map(grado_map)

    filasn, _ = df_nombres.shape

    df_bloque = pd.DataFrame([""] * filasn)
    df_asignatura = pd.DataFrame([""] * filasn)
    df_desempeño1 = pd.DataFrame([""] * filasn)
    df_desempeño2 = pd.DataFrame([""] * filasn)
    df_desempeño3 = pd.DataFrame([""] * filasn)
    df_desempeño4 = pd.DataFrame([""] * filasn)
    df_desempeño5 = pd.DataFrame([""] * filasn)

    if (a == 'M1') or (a =='M2'):
        
        for i in range(filasn):
            estudiante_actual = df_nombres.iloc[i, 0]
            grado_actual = df_nombres.iloc[i, 2]
            notas_estudiante = dfn[(dfn.iloc[:, 2] == estudiante_actual) & (dfn.iloc[:, 3] == grado_actual)]
            desempeno_encontrado = False
        
            if df_nombres.iloc[i, 0] in (['LMOD1', 'LMOD2', 'MMOD1', 'MMOD2', 'WMOD1', 'WMOD2', 'JMOD1', 'JMOD2', 'VMOD1']):
                desempeno_encontrado = True
                continue 
            
            materias_especificas_6_7 = ['Aritmética', 'Geometría', 'Estadística', 'Dibujo técnico', 'Sistemas']
            materias_especificas_8_9 = ['Álgebra', 'Geometría', 'Estadística', 'Dibujo técnico', 'Sistemas']
            materias_especificas_10 =  ['Trigonometría', 'Matemática financiera', 'Estadística', 'Dibujo técnico', 'Sistemas']
            materias_especificas_11 =  ['Cálculo', 'Matemática financiera', 'Estadística', 'Dibujo técnico', 'Sistemas']
                
            bloques = ['A', 'B', 'C', 'D']
            asignaturas = ['Aritmética', 'Geometría','Matemática financiera', 'Álgebra', 'Estadística', 'Dibujo técnico', 'Sistemas','Trigonometría','Cálculo']

            for materia in asignaturas:
                notas_materia = notas_estudiante[notas_estudiante.iloc[:, 5] == materia]
                
                for bloque in bloques:
                    notas_bloque = notas_materia[notas_materia.iloc[:, 6] == bloque]
                    desempenos_completos = len(notas_bloque)

                    if desempenos_completos == 1:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño2.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 2:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño3.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 3:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño4.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 4: 
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño5.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                if desempeno_encontrado:
                    break
                
            for bloque in bloques:
                notas_bloque = notas_estudiante[notas_estudiante.iloc[:, 6] == bloque]
                if grado_actual in (6,7):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_6_7)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 25 and not desempeno_encontrado:
                        continue 
                    if longitud_bloque < 25 and not desempeno_encontrado:
                        for materia in materias_especificas_6_7:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual in (8,9):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_8_9)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 25 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 25 and not desempeno_encontrado:
                        for materia in materias_especificas_8_9:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual == 10:
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_10)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 25 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 25 and not desempeno_encontrado:
                        for materia in materias_especificas_10:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual == 11:
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_11)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 25 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 25 and not desempeno_encontrado:
                        for materia in materias_especificas_11:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if desempeno_encontrado:
                    break

    if  (a == 'C1') or (a =='C2'):
        
        for i in range(filasn):
            
            estudiante_actual = df_nombres.iloc[i, 0]
            grado_actual = df_nombres.iloc[i, 2]
            notas_estudiante = dfn[(dfn.iloc[:, 2] == estudiante_actual) & (dfn.iloc[:, 3] == grado_actual)]
            desempeno_encontrado = False
        
            if df_nombres.iloc[i, 0] in (['LMOD1', 'LMOD2', 'MMOD1', 'MMOD2', 'WMOD1', 'WMOD2', 'JMOD1', 'JMOD2', 'VMOD1']):
                desempeno_encontrado = True
                continue 
            
                
            bloques = ['A', 'B', 'C', 'D']
            asignaturas = ['Biología', 'Química', 'Medio ambiente', 'Física']
            asignaturas_11 = ['Química', 'Medio ambiente', 'Física']

            for materia in asignaturas:
                notas_materia = notas_estudiante[notas_estudiante.iloc[:, 5] == materia]
                
                for bloque in bloques:
                    notas_bloque = notas_materia[notas_materia.iloc[:, 6] == bloque]
                    desempenos_completos = len(notas_bloque)

                    if desempenos_completos == 1:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño2.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 2:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño3.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 3:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño4.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 4: 
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño5.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                if desempeno_encontrado:
                    break
                
            for bloque in bloques:
                
                notas_bloque = notas_estudiante[notas_estudiante.iloc[:, 6] == bloque]
                
                if grado_actual in (6,7,8,9,10):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(asignaturas)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 20 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 20 and not desempeno_encontrado:
                        for materia in asignaturas:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual == 11:
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(asignaturas_11)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 15 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 15 and not desempeno_encontrado:
                        for materia in asignaturas_11:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if desempeno_encontrado:
                    break
                

    if a == 'S1':
        
        for i in range(filasn):
            estudiante_actual = df_nombres.iloc[i, 0]
            grado_actual = df_nombres.iloc[i, 2]
            notas_estudiante = dfn[(dfn.iloc[:, 2] == estudiante_actual) & (dfn.iloc[:, 3] == grado_actual)]
            desempeno_encontrado = False
        
            if df_nombres.iloc[i, 0] in (['LMOD1', 'LMOD2', 'MMOD1', 'MMOD2', 'WMOD1', 'WMOD2', 'JMOD1', 'JMOD2', 'VMOD1']):
                desempeno_encontrado = True
                continue 
            
            materias_especificas_6_7_8_9 = ['Historia', 'Geografía', 'Participación política', 'Filosofía']
            materias_especificas_10_11 = ['Ciencias económicas', 'Ciencias políticas', 'Filosofía']
        
                    
            bloques = ['A', 'B', 'C', 'D']
            asignaturas = ['Historia', 'Geografía', 'Ciencias económicas', 'Participación política','Ciencias políticas','Filosofía']

            for materia in asignaturas:
                notas_materia = notas_estudiante[notas_estudiante.iloc[:, 5] == materia]
                
                for bloque in bloques:
                    notas_bloque = notas_materia[notas_materia.iloc[:, 6] == bloque]
                    desempenos_completos = len(notas_bloque)

                    if desempenos_completos == 1:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño2.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 2:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño3.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 3:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño4.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 4: 
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño5.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                if desempeno_encontrado:
                    break
                
            for bloque in bloques:
                notas_bloque = notas_estudiante[notas_estudiante.iloc[:, 6] == bloque]
                
                
                if grado_actual in (6,7,8,9):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_6_7_8_9)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 20 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 20 and not desempeno_encontrado:
                        for materia in materias_especificas_6_7_8_9:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual in (10,11):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_10_11)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 15 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 15 and not desempeno_encontrado:
                        for materia in materias_especificas_10_11:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if desempeno_encontrado:
                    break


    if a == 'S2':
        
        for i in range(filasn):
            estudiante_actual = df_nombres.iloc[i, 0]
            grado_actual = df_nombres.iloc[i, 2]
            notas_estudiante = dfn[(dfn.iloc[:, 2] == estudiante_actual) & (dfn.iloc[:, 3] == grado_actual)]
            desempeno_encontrado = False
        
            if df_nombres.iloc[i, 0] in (['LMOD1', 'LMOD2', 'MMOD1', 'MMOD2', 'WMOD1', 'WMOD2', 'JMOD1', 'JMOD2', 'VMOD1']):
                desempeno_encontrado = True
                continue 
            
            materias_especificas_6_7_8_9 = ['Historia', 'Geografía', 'Participación política', 'Filosofía']
            materias_especificas_10_11 = ['Ciencias económicas', 'Ciencias políticas', 'Filosofía']
        
                    
            bloques = ['A', 'B', 'C', 'D']
            asignaturas = ['Historia', 'Geografía', 'Ciencias económicas', 'Participación política','Ciencias políticas','Filosofía']

            for materia in asignaturas:
                notas_materia = notas_estudiante[notas_estudiante.iloc[:, 5] == materia]
                
                for bloque in bloques:
                    notas_bloque = notas_materia[notas_materia.iloc[:, 6] == bloque]
                    desempenos_completos = len(notas_bloque)

                    if desempenos_completos == 1:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño2.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 2:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño3.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 3:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño4.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 4: 
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño5.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                if desempeno_encontrado:
                    break
                
            for bloque in bloques:
                notas_bloque = notas_estudiante[notas_estudiante.iloc[:, 6] == bloque]
                
                
                if grado_actual in (6,7,8,9):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_6_7_8_9)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 20 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 20 and not desempeno_encontrado:
                        for materia in materias_especificas_6_7_8_9:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if grado_actual in (10,11):
                    notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(materias_especificas_10_11)]
                    longitud_bloque = len(notas_bloque_filtradas)
                    if longitud_bloque == 20 and not desempeno_encontrado:
                        continue 
        
                    if longitud_bloque < 20 and not desempeno_encontrado:
                        for materia in materias_especificas_10_11:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if desempeno_encontrado:
                    break


    if a == 'L':
        
        for i in range(filasn):
            estudiante_actual = df_nombres.iloc[i, 0]
            grado_actual = df_nombres.iloc[i, 2]
            notas_estudiante = dfn[(dfn.iloc[:, 2] == estudiante_actual) & (dfn.iloc[:, 3] == grado_actual)]
            desempeno_encontrado = False
        
            if df_nombres.iloc[i, 0] in (['LMOD1', 'LMOD2', 'MMOD1', 'MMOD2', 'WMOD1', 'WMOD2', 'JMOD1', 'JMOD2', 'VMOD1']):
                desempeno_encontrado = True
                continue 
            
                
            bloques = ['A', 'B', 'C', 'D']
            asignaturas = ['Comunicación y sistemas simbólicos', 'Producción e interpretación de textos','Metodología']

            for materia in asignaturas:
                notas_materia = notas_estudiante[notas_estudiante.iloc[:, 5] == materia]
                
                for bloque in bloques:
                    notas_bloque = notas_materia[notas_materia.iloc[:, 6] == bloque]
                    desempenos_completos = len(notas_bloque)

                    if desempenos_completos == 1:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño2.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 2:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño3.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 3:
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño4.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                    elif desempenos_completos == 4: 
                        df_bloque.iloc[i] = bloque
                        df_asignatura.iloc[i] = materia
                        df_desempeño5.iloc[i] = 'X'
                        desempeno_encontrado = True
                        break

                if desempeno_encontrado:
                    break
                
            for bloque in bloques:
                
                notas_bloque = notas_estudiante[notas_estudiante.iloc[:, 6] == bloque]
                notas_bloque_filtradas = notas_bloque[notas_bloque.iloc[:, 5].isin(asignaturas)]
                longitud_bloque = len(notas_bloque_filtradas)
                if longitud_bloque == 10 and not desempeno_encontrado:
                    continue 
        
                if longitud_bloque < 10 and not desempeno_encontrado:
                    for materia in asignaturas:
                            if materia not in notas_bloque_filtradas.iloc[:, 5].values:
                                df_bloque.iloc[i] = bloque
                                df_desempeño1.iloc[i] = 'X'
                                df_asignatura.iloc[i] = materia
                                desempeno_encontrado = True
                                break
                                
                if desempeno_encontrado:
                    break
    
    df_horario = pd.concat([df_nombres, df_bloque, df_asignatura, df_desempeño1, df_desempeño2, df_desempeño3, df_desempeño4, df_desempeño5], ignore_index=True, axis=1)

    for i in range(len(df_horario)):
        for k in range(i + 1, len(df_horario)):
            if df_horario.iloc[i, 0] == df_horario.iloc[k, 0]:
                if 'X' in df_horario.iloc[i].values:
                    b = df_horario.iloc[i].eq('X').idxmax()
                    if b + 1 < len(df_horario.columns):
                        df_horario.iloc[k, b + 1] = 'X'
                        for col in range(5, b + 1):
                            df_horario.iloc[k, col] = np.nan
                break
                
    if a == 'M1':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSM1.xlsx', index=False)
        print('F1.2 Matematicas 1 hecho')
    if a == 'M2':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSM2.xlsx', index=False)
        print('F1.2 Matematicas 2 hecho')
    if a == 'C1':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSC1.xlsx', index=False)
        print('F1.2 Ciencias 1 hecho')
    if a == 'C2':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSC2.xlsx', index=False)
        print('F1.2 ciencias 2 hecho')
    if a == 'S1':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSS1.xlsx', index=False)
        print('F1.2 Sociales 1 hecho')
    if a == 'S2':
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSS2.xlsx', index=False)
        print('F1.2 Sociales 2 hecho')
    if a == 'L':
        df_horario[4] = df_horario[4].replace('Comunicación y sistemas simbólicos','Com')
        df_horario[4] = df_horario[4].replace('Producción e interpretación de textos','Pro')
        df_horario.to_excel('C:/Users/Admin/Desktop/GK/F1.2/Bachillerato/HSL.xlsx', index=False)
        print('F1.2 Lenguaje hecho')