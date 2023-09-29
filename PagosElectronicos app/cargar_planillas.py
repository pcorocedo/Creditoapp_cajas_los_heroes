import logging
import os
import sys
import pandas as pd
import pyodbc
from cryptography.fernet import Fernet
from openpyxl import load_workbook
import datetime
from pandas.tseries.offsets import BDay
import locale
locale.setlocale(locale.LC_ALL, "")
logging.basicConfig(filename="data.log", level=logging.INFO,
                    format='%(asctime)s :: %(levelname)s :: %(funcName)s :: %(lineno)d :: %(message)s')
def definir_feriados_siguiente(today):
    global dia_habil_siguiente2
    try:
        server = 'bansqlcl'
        database = 'Produccion_BT'

        username = 'botcredito'
        password = 'OMA7x5qk'
        conn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server +
            ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)

        cursor = conn.cursor()
        cursor.execute(
            "select CONVERT(varchar,Ffecha,105) as Ffecha ,fhabil from FST028 (nolock) where fhabil ='N' order by ffecha ASC")
        lista_feriados = []
        for row in cursor:
            lista_feriados.append(row[0])
        dia_habil_siguiente2 = datetime.datetime.strptime(str(today), "%Y%m%d")
        encontrado = 0
        habil = 'NO'
        if datetime.datetime.strftime(dia_habil_siguiente2, '%A') != 'sabado' and datetime.datetime.strftime(dia_habil_siguiente2,'%A') != 'domingo' and datetime.datetime.strftime(dia_habil_siguiente2, "%d-%m-%Y") not in lista_feriados:
            dia_habil_siguiente2 = str(today)
            habil ='SI'
        else:
            for i in range(0, 20):
                dia_habil_siguiente2 = str(dia_habil_siguiente2 + BDay(1))[:10]
                dia_habil_siguiente2 = datetime.datetime.strptime(str(dia_habil_siguiente2), "%Y-%m-%d")
                print(datetime.datetime.strftime(dia_habil_siguiente2, '%A'))
                if datetime.datetime.strftime(dia_habil_siguiente2, '%A') != 'sabado' and datetime.datetime.strftime(dia_habil_siguiente2,
                                                                                               '%A') != 'domingo':
                    dia_habil_siguiente2 = datetime.datetime.strftime(dia_habil_siguiente2, "%d-%m-%Y")
                    if dia_habil_siguiente2 not in lista_feriados:
                        encontrado = encontrado + 1
                        if encontrado == 1:
                            break
                    dia_habil_siguiente2 = datetime.datetime.strptime(str(dia_habil_siguiente2), "%d-%m-%Y")
        if habil =='NO':
            dia_habil_siguiente2 = str(dia_habil_siguiente2[len(dia_habil_siguiente2)-4:]) + str(dia_habil_siguiente2[len(dia_habil_siguiente2)-7:len(dia_habil_siguiente2)-5]) + str(dia_habil_siguiente2[:2])
    except Exception as ftperror:
        logging.error(ftperror.args[1])
        sys.exit()
def cuadrar_plac(lineas,sftp_path,archivo_ac):
    server = server_bd
    database = base_datos

    username = user_bd
    password = pass_bd
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server +
        ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)

    cursor = conn.cursor()
    lista_folios_ac = list(filter(lambda filtro_folios: filtro_folios[0] == "2", lineas))
    monto_ab_total = 0
    totalpl = 0
    for folios_ac in lista_folios_ac:

        lista_datos = []
        folio = folios_ac[1:17]
        banco = folios_ac[31:47].strip()
        monto_ab = folios_ac[170:182]
        monto_ab = monto_ab.lstrip('0')
        lista_datos.append(folio)
        lista_datos.append(banco)
        monto_ab_total = monto_ab_total + int(monto_ab)

        # Obtener datos desde archivo PL

        archivo_pl = archivo_ac.replace('AC', 'PL')
        f2 = open(sftp_path + "/" + archivo_pl, "r")
        lineas2 = f2.readlines()
        f2.close()
        folio_encontrado = False
        for l2 in lineas2:
            if folio in l2:
                folio_encontrado = True
            if folio_encontrado and l2[0] == '4':
                cuerpo4_cols = get_columnas(l2)
                lista_datos.extend(cuerpo4_cols)
                break
        totalpl = totalpl + int(cuerpo4_cols[68])
    if monto_ab_total==totalpl:
        cuadratura_planillas ='OK'
    else:
        cuadratura_planillas = 'NOK'
    valores = archivo_pl,str(totalpl),archivo_ac,monto_ab_total,cuadratura_planillas
    cursor.execute("insert into cuadratura_planillas (archivo_pl,monto_pl,archivo_ac,monto_ac,estado_cuadratura) values" + str(valores) + " ")
    cursor.commit()
    return cuadratura_planillas
def cargar_llave():
    return open("llave.key", "rb").read()
def acceso():
    global server_bd,user_bd,pass_bd,base_datos
    try:
        f = Fernet(cargar_llave())
        excel2 = os.path.join(os.getcwd(), "acceso.xlsx")
        wb2 = load_workbook(excel2)
        ws_claves = wb2['claves']
        server_bd = ws_claves["D4"].value
        server_bd = (f.decrypt(server_bd.encode())).decode()
        user_bd = ws_claves["B4"].value
        user_bd = (f.decrypt(user_bd.encode())).decode()
        pass_bd = ws_claves["C4"].value
        pass_bd = (f.decrypt(pass_bd.encode())).decode()
        base_datos = ws_claves["E4"].value
        base_datos = (f.decrypt(base_datos.encode())).decode()
    except Exception as ftperror:
        logging.error(ftperror.args[1])
        sys.exit()
def get_columnas(line: str) -> list:
    """
    Genera lista con las columnas separadas desde un string
    :param line: string
    :return: list
    """

    nro_trab_h_no_afi = line[1:13]
    nro_trab_m_no_afi = line[13:25]
    rem_no_afi = line[25:37]
    cot_no_afi = line[37:49]
    nro_trab_h_afi = line[49:61]
    nro_trab_m_afi = line[61:73]
    rem_afi = line[73:85]
    cot_afi_isapre = line[85:97]
    nro_trab_aporte_1 = line[97:109]
    total_rem_trab_aporte_1 = line[109:121]
    cot_aporte_1 = line[121:133]
    nro_trab_aporte_adic = line[133:145]
    total_rem_trab_aporte_adic = line[145:157]
    cot_aporte_adic = line[157:169]
    nro_trab_cred_personales = line[169:181]
    total_rem_trab_cred_personales = line[181:193]
    cot_cred_personales = int(line[193:205])
    nro_trab_conv_dental = line[205:217]
    total_rem_trab_conv_dental = line[217:229]
    con_conv_dental = int(line[229:241])
    nro_trab_leasing = line[241:253]
    total_rem_trab_leasing = line[253:265]
    cot_leasing = int(line[265:277])
    nro_trab_seg_vida = line[277:289]
    total_rem_trab_seg_vida = line[289:301]
    cot_seg_vida = int(line[301:313])
    nro_trab_otros_dctos = line[313:325]
    total_rem_trab_otros_dctos = line[325:337]
    cot_otros_dctos = int(line[337:349])
    total_rem = line[349:361]
    total_cot = line[361:373]
    nro_cargas_A_simples = line[373:385]
    nro_cargas_A_invalidez = line[385:397]
    nro_cargas_A_maternales = line[397:409]
    cant_trab_A = line[409:421]
    monto_rebaja_A = line[421:433]
    nro_cargas_B_simples = line[433:445]
    nro_cargas_B_invalidez = line[445:457]
    nro_cargas_B_maternales = line[457:469]
    cant_trab_B = line[469:481]
    monto_rebaja_B = line[481:493]
    nro_cargas_C_simples = line[493:505]
    nro_cargas_C_invalidez = line[505:517]
    nro_cargas_C_maternales = line[517:529]
    cant_trab_C = line[529:541]
    monto_rebaja_C = line[541:553]
    nro_cargas_D_simples = line[553:565]
    nro_cargas_D_invalidez = line[565:577]
    nro_cargas_D_maternales = line[577:589]
    cant_trab_D = line[589:601]
    total_cargas_simples = line[601:613]
    total_cargas_invalidas = line[613:625]
    total_cargas_maternales = line[625:637]
    total_asignacion_familiar = line[637:649]
    monto_rebaja_asig_fam_retro = line[649:661]
    reintegros_asig_fam = line[661:673]
    total_rebajas = line[673:685]
    saldo = line[685:697]
    total_rem_60 = line[697:709]
    total_gratificaciones = line[709:721]
    periodo_desde = line[721:727]
    periodo_hasta = line[727:733]
    nro_afi_informados = line[733:740]
    fecha_pago = line[740:748]
    periodo_rem = line[748:754]
    prod_fin = int(cot_cred_personales) + int(con_conv_dental) + int(cot_leasing) + int(cot_seg_vida) + int(cot_otros_dctos)
    afa = int(cot_no_afi) - int(total_rebajas)
    unoporc = int(cot_aporte_1)
    if afa < 0:
        suma_pl = prod_fin + unoporc
        afa = 0
    else:
        suma_pl = prod_fin + afa + unoporc

    columnas = [nro_trab_h_no_afi, nro_trab_m_no_afi, rem_no_afi, cot_no_afi, nro_trab_h_afi, nro_trab_m_afi, rem_afi,
                cot_afi_isapre, nro_trab_aporte_1, total_rem_trab_aporte_1, cot_aporte_1, nro_trab_aporte_adic,
                total_rem_trab_aporte_adic, cot_aporte_adic, nro_trab_cred_personales, total_rem_trab_cred_personales,
                cot_cred_personales, nro_trab_conv_dental, total_rem_trab_conv_dental, con_conv_dental,
                nro_trab_leasing, total_rem_trab_leasing, cot_leasing, nro_trab_seg_vida, total_rem_trab_seg_vida,
                cot_seg_vida, nro_trab_otros_dctos, total_rem_trab_otros_dctos, cot_otros_dctos, total_rem, total_cot,
                nro_cargas_A_simples, nro_cargas_A_invalidez, nro_cargas_A_maternales, cant_trab_A, monto_rebaja_A,
                nro_cargas_B_simples, nro_cargas_B_invalidez, nro_cargas_B_maternales, cant_trab_B, monto_rebaja_B,
                nro_cargas_C_simples, nro_cargas_C_invalidez, nro_cargas_C_maternales, cant_trab_C, monto_rebaja_C,
                nro_cargas_D_simples, nro_cargas_D_invalidez, nro_cargas_D_maternales, cant_trab_D,
                total_cargas_simples, total_cargas_invalidas, total_cargas_maternales, total_asignacion_familiar,
                monto_rebaja_asig_fam_retro, reintegros_asig_fam, total_rebajas, saldo, total_rem_60,
                total_gratificaciones, periodo_desde, periodo_hasta, nro_afi_informados, fecha_pago, periodo_rem,prod_fin, afa, unoporc,suma_pl]

    return columnas
def get_datos(lista_archivos: list, sftp_path: str) -> pd.DataFrame:
    """
    Obtiene datos de archivos almacenados localmente y guarda la información en una base de datos
    """
    server = server_bd
    database = base_datos

    username = user_bd
    password = pass_bd
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server +
        ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)

    cursor = conn.cursor()
    lista_ac = list(filter(lambda filtro_lista: filtro_lista[:2] == "AC", lista_archivos))

    df_datos = pd.DataFrame(
        columns=['Folio', 'Banco', 'Número de Trabajadores Hombres no Afiliados',
                 'Número de Trabajadores Mujeres no Afiliados', 'Remuneraciones no Afiliados',
                 'Cotización no Afiliados', 'Número de Trabajadores Hombres Afiliados',
                 'Número de Trabajadores Mujeres Afiliados', 'Remuneraciones Afiliados',
                 'Cotización de Afiliados a Isapres', 'Numero de Trabajadores Aporte 1%',
                 'Total Remuneraciones Trabajadores Aporte 1%', 'Cotización por Aporte 1%',
                 'Número de Trabajadores Aporte Adicional', 'Total Remuneraciones Trabajadores Aporte Adicional',
                 'Cotización por Aporte Adicional', 'Número Trabajadores Creditos Personales',
                 'Total Remuneraciones Trabajadores Creditos Personales', 'Cotización por Créditos Personales',
                 'Número de Trabajadores con Convenio Dental',
                 'Total de Remuneraciones de Trabajadores con Convenio Dental', 'Cotización Convenio Dental',
                 'Número Trabajadores Leasing', 'Total Remuneraciones Trabajadores Leasing', 'Cotización Leasing',
                 'Número de Trabajadores con Seguros de Vida',
                 'Total de remuneraciones de Trabajadores con Seguros de Vida', 'Cotización Seguros de Vida',
                 'Número de Trabajadores con Otros Descuentos', 'Total de Remuneraciones con Otros Descuentos',
                 'Cotización Otros Descuentos', 'Total remuneraciones', 'Total Cotizaciones',
                 'Número de Cargas Tramo A Simples', 'Número de Cargas Tramo A Invalidez',
                 'Número de Cargas Tramo A Maternales', 'Cantidad de Trabajadores Tramo A', 'Monto Rebaja Tramo A',
                 'Número de Cargas Tramo B Simples', 'Número de Cargas Tramo B Invalidez',
                 'Número de Cargas Tramo B Maternales', 'Cantidad de Trabajadores Tramo B', 'Monto Rebaja Tramo B',
                 'Número de Cargas Tramo C Simples', 'Número de Cargas Tramo C Invalidez',
                 'Número de Cargas Tramo C Maternales', 'Cantidad de Trabajadores Tramo C', 'Monto Rebaja Tramo C',
                 'Número de Cargas Tramo D Simple', 'Número de Cargas Tramo D Invalidez',
                 'Número de Cargas Tramo D Maternales', 'Cantidad de Trabajadores Tramo D', 'Total de Cargas Simples',
                 'Total de Cargas Invalidas', 'Total de Cargas Maternales', 'Total de Asignación Familiar',
                 'Monto de Rebaja Asignación Familiar Retro.', 'Reintegros Asignación Familiar', 'Total Rebajas',
                 'Saldo', 'Total Remuneraciones 60', 'Total Gratificaciones', 'Período desde', 'Período hasta ',
                 'Número de afiliados informados', 'Fecha de pago', 'Periodo de Remuneración','prod_fin', 'afa', 'unoporc','suma_pl','cta_cte','fecha_abono','fecha_operacion'])
    for archivo_ac in lista_ac:
        try:
            f = open(sftp_path + "/" + archivo_ac, "r")
            lineas = f.readlines()
            f.close()
            resultado_cuadratura = cuadrar_plac(lineas,sftp_path,archivo_ac)
            if resultado_cuadratura == 'OK':
                lista_folios_ac = list(filter(lambda filtro_folios: filtro_folios[0] == "2", lineas))
                for folios_ac in lista_folios_ac:

                    lista_datos = []
                    fecha_ab = folios_ac[87:95]
                    definir_feriados_siguiente(fecha_ab)
                    fecha_ab = dia_habil_siguiente2
                    folio = folios_ac[1:17]
                    banco = folios_ac[31:47].strip()
                    fecha_op = folios_ac[17:25]
                    rut_pagador = folios_ac[123:132]
                    rut_pagador = rut_pagador.lstrip('0')
                    dv_pagador = folios_ac[132:133]
                    monto_ab = folios_ac[170:182]
                    monto_ab = monto_ab.lstrip('0')
                    nro_cta_cte = folios_ac[182:212].strip()
                    nro_cta_cte = nro_cta_cte.lstrip('0')
                    lista_datos.append(folio)
                    lista_datos.append(banco)


                    # Obtener datos desde archivo PL

                    archivo_pl = archivo_ac.replace('AC', 'PL')
                    f2 = open(sftp_path + "/" + archivo_pl, "r")
                    lineas2 = f2.readlines()
                    f2.close()
                    folio_encontrado = False
                    for l2 in lineas2:
                        if folio in l2:
                            folio_encontrado = True
                        if folio_encontrado and l2[0] == '4':
                                cuerpo4_cols = get_columnas(l2)
                                lista_datos.extend(cuerpo4_cols)
                                break
                    lista_datos.append(nro_cta_cte)
                    lista_datos.append(fecha_ab)
                    lista_datos.append(fecha_op)
                    df_datos.loc[len(df_datos)] = lista_datos
                    diferencia = int(lista_datos[70]) - int(monto_ab)
                    if int(diferencia) ==0:
                        estado='OK'
                    else: estado='NOK'
                    valores= archivo_ac,folio,rut_pagador + '-' + str(dv_pagador) ,banco,str(lista_datos[18]),str(lista_datos[21]),str(lista_datos[24]),str(lista_datos[27]),str(lista_datos[30]),str(lista_datos[67]),str(lista_datos[68]),str(lista_datos[69]),str(lista_datos[70]),str(monto_ab),str(diferencia),estado,fecha_op,fecha_ab,nro_cta_cte
                    cursor.execute("insert into planillas (nombre_archivo,folio,rut,banco,credito_personal,convenio_dental,leasing,seguro_vida,otros_descuentos,prod_fin,afa,unoporciento,suma_pl,monto_ac,diferencia,estado_cuadra,fecha_operacion,fecha_abono,cta_cte) values" + str(valores) + " ")
        except:
            valores = archivo_ac, folio,"ERROR ARCHIVO PL"
            cursor.execute("insert into planillas (nombre_archivo,folio,estado_cuadra) values" + str(valores) + " ")

    cursor.commit()
    return df_datos
def buscar_data():
    try:
        server = server_bd
        database = base_datos

        username = user_bd
        password = pass_bd
        conn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server +
            ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)

        cursor = conn.cursor()
        cursor.execute("select nombre_archivo from planillas")
        lista =  [item[0] for item in cursor.fetchall()]
        path = os.path.join(os.getcwd(), "planillas")
        filelist = os.listdir(path)
        for recorre in reversed(lista):
            try:
                result = filelist.index(recorre)
                filelist.pop(result)
            except:
                pass
        get_datos(filelist, path)
    except Exception as ftperror:
        logging.error(ftperror.args[1])
        sys.exit()
acceso()
buscar_data()
