def generar_turnosJavelin(request):
    if request.user.is_authenticated and request.user.has_perm('solicitudes.read_detallesolicitud'):
        detalle = {}
        results = []
        today = datetime.date.today()
        cliente = ''
        sucursal = ''
        desde = today
        hasta= today
        tipo_solicitud = ''
        estado =''
        solicitudes = []
        import re
        ##obtener los clientes
        connection = pyhdb.connect(
                host=P_SERVER,
                port=P_PORT,
                user=P_USER_SERVER,
                password=P_PASSWD_SERVER
            )
        cursor = connection.cursor()
        sql = """SELECT "CardCode","CardName" FROM "%s"."OCRD" T0 WHERE T0."CardType" in ('C') and T0."validFor"='Y' """ % (P_COMPANY)
        cursor.execute(sql)
        listado_clientes = cursor.fetchall()
        dict_clientes = {}
        for p in listado_clientes:
            dict_clientes['%s' % p[0]] = p[1]


        sql = """SELECT "CODEMPLEADO","CARGO" FROM "%s"."MAESTROS_EMPLEADO" """ % (ESQUEMA_VENTAS)
        cursor.execute(sql)
        listado_empleados = cursor.fetchall()
        dict_empleados = {}
        for p in listado_empleados:
            dict_empleados['%s' % p[0]] = p[1]
        try:
            connection.close()
        except:
            pass

        if request.method == 'POST':
            tipo_solicitud = 'SSEE'
            cliente = request.POST['cliente']
            sucursal = request.POST['sucursal']    

            desde = datetime.datetime.strptime(request.POST['desde'],'%d/%m/%Y')
            hasta = datetime.datetime.strptime(request.POST['hasta'],'%d/%m/%Y')

            fecha_desde_solo = desde.strftime('%Y-%m-%d')
            fecha_hasta_solo = hasta.strftime('%Y-%m-%d')
            estado = '3'

            ### VALIDACION ###
            connection = MySQLdb.connect(host=P_SERVER_JC,user=P_USER_JC,passwd=P_CLAVE_JC, port=int(P_PORT_JC))
            mysql_cursor = connection.cursor()
            sqlval = """
            select distinct
            emp.legacy_employee_identifier AS id_employee,
            sch.quick_shift_desc as cod,
            date(convert_tz(sch.shift_start_date_time,'Etc/UTC','America/Lima')) as initial_date
            from  g4s_scheduler.schedule_history sch
            left join employees_and_assets.employee emp on emp.id = sch.employee_id
            left join customers_and_contracts.WORKORDER_LOCATION wol ON wol.id = sch.work_order_location_id
            left join customers_and_contracts.WORK_ORDER wor ON wor.id = wol.work_order_id
            left join customers_and_contracts.contract ctr ON ctr.id = wor.contract_id
            left join customers_and_contracts.account acc ON acc.id = ctr.account_id
            left join customers_and_contracts.customer cus ON cus.id = acc.customer_id
            left join customers_and_contracts.billing_information bil on wor.billing_information_id=bil.id
            left join customers_and_contracts.shift_details sh
            on sh.id=sch.shift_id and sch.post_id=sh.post_id
            where acc.customerReference1 is not null
            and date(convert_tz(sch.shift_start_date_time,'Etc/UTC','America/Lima'))
            between '%s' and '%s'
            """% (
            fecha_desde_solo,
            fecha_hasta_solo)  
            mysql_cursor.execute(sqlval)
            result_dict = {}

            for row in mysql_cursor.fetchall():
                id_employee,cod,initial_date,  = row
                formatted_date = initial_date.strftime('%Y-%m-%d')
                key = (id_employee,cod,formatted_date)
                result_dict[key] = {
                    'id_employee': id_employee,
                    'cod': cod,
                    'initial_date': formatted_date,
                }
            
            print(result_dict)
            mysql_cursor.close()
            connection.close()
            

            ### CONEXION MYSQL ###
            connection = MySQLdb.connect(host=P_SERVER_JC,user=P_USER_JC,passwd=P_CLAVE_JC, port=int(P_PORT_JC))
            mysql_cursor = connection.cursor()
            sql = """
            SELECT
            pue.id,
            pue.name AS name,
            s.addressLine3
            -- s.siteName AS postnamessee
            FROM customers_and_contracts.WORKORDER_LOCATION wol
            LEFT JOIN customers_and_contracts.POST pue ON wol.id = pue.work_order_location_id
            LEFT JOIN customers_and_contracts.contract_site cs ON cs.id = wol.contract_site_id
            LEFT JOIN customers_and_contracts.site s ON s.id = cs.site_id
            LEFT JOIN customers_and_contracts.WORK_ORDER wo ON wol.work_order_id = wo.id
            LEFT JOIN customers_and_contracts.billing_information bi ON bi.id = wo.billing_information_id
            LEFT JOIN customers_and_contracts.contract co ON co.id = wo.contract_id
            LEFT JOIN customers_and_contracts.account a ON a.id = co.account_id
            LEFT JOIN customers_and_contracts.customer c ON c.id = a.customer_id
            LEFT JOIN customers_and_contracts.shift_details sh ON sh.post_id = pue.id
            LEFT JOIN customers_and_contracts.shift_details sh1
            ON sh.post_id = sh1.post_id AND sh.ammendment_no = sh1.ammendment_no - 1
            WHERE 1=1
            AND sh1.id IS NULL
            AND a.customerReference1 IS NOT NULL
            AND wol.status = 'ACTIVE'
            AND bi.work_order_type_id = 62
            """

            mysql_cursor.execute(sql)
            location_name_mysql = mysql_cursor.fetchall()
            # Almacena los resultados en un diccionario
            data_dict = {}
            for fila in location_name_mysql:
                id_pue = fila[0]
                nombre_pue = fila[1]
                direccion = fila[2]

                data_dict[id_pue] = {
                    "nombre": nombre_pue,
                    "direccion": direccion
                }

            mysql_cursor.close()
            connection.close()

            def buscar_id_pue_por_id_unidad(data_dict, id_unidad):
                for id_pue, detalles in data_dict.items():
                    if detalles["direccion"] == id_unidad:
                        return id_pue
                return None
            
            if desde != '' and hasta != '':
                solicitudes = MVC_Solicitud_Servicio.objects.filter(fechaInicio__gte=desde,fechaFin__lte=hasta)
            if cliente != '':
                solicitudes = solicitudes.filter(idSolicitud__codCliente=cliente)
            if sucursal != '':
                qs = solicitudes.filter(codBranch = None)
                solicitudes = solicitudes.filter(codBranch_id=sucursal)
            if estado != '':
                estado = MVC_Lv_Estado_Solicitud_Servicio.objects.filter(idEstado = estado)
                if estado.count() > 0:
                    solicitudes = solicitudes.filter(idEstado=estado[0].idEstado)
        else:
            solicitudes = MVC_Solicitud_Servicio.objects.filter(fechaInicio__gte=today,fechaFin__lte=today)
            
        regis_puesto = {}
        for s in solicitudes:
            for e in s.g4s_estado_solicitud_set.all():
                if e.Soline in regis_puesto.keys():
                    pass
                else:
                    if e.Soline:
                        posDet = MVC_Os_Posicion_Det.objects.filter(idOSPosicionDet = int(e.Soline))
                        if posDet.count() > 0:
                            if posDet[0].idListaPrecios:
                                regis_puesto['%s' % e.Soline] = {'codigosaturn1':'','codigosaturn2':''}
                                regis_puesto['%s' % e.Soline]['codigoSaturn1'] = posDet[0].codigoSaturn1
                                regis_puesto['%s' % e.Soline]['codigoSaturn2'] = posDet[0].codigoSaturn2
        #embacezado
        results.append(['','','YYYY-MM-DD','YYYY-MM-DD','HH:MM','HH:MM','','Y/N','Y/N','','','0/1','0/1',''])
        results.append(['Legacy Site ID','Person ID','Start Date','End Date','Start Time','End Time','Time Zone','Check In','Check Out','Check Call Frequency','Check Call Time','Suppress First Check Call','Suppress Last Check Call','Description'])
        #llenado de data
        if tipo_solicitud == '' or tipo_solicitud == T_SERVICIO:
            for s in solicitudes:
                for e in s.g4s_estado_solicitud_set.all():
                    detalle = []
                    id_unidad=s.idAssignment.pk
                    resultado = buscar_id_pue_por_id_unidad(data_dict, str(id_unidad))
                    try:
                        detalle.append(resultado)
                    except:
                        detalle.append('')
                    if e.Codempleado is None:
                        detalle.append("")
                    else:
                        detalle.append(str(e.Codempleado)[:8].zfill(10))
                    try:
                        detalle.append(s.fechaInicio.strftime('%Y-%m-%d'))
                    except:
                        detalle.append('')
                    try:
                        detalle.append(s.fechaFin.strftime('%Y-%m-%d'))
                    except:
                        detalle.append('')
                    try:
                        detalle.append(s.horaInicio.strftime('%H:%M'))
                    except:
                        detalle.append('')
                    try:
                        detalle.append(s.horaFin.strftime('%H:%M'))
                    except:
                        detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    detalle.append('')
                    puesto_posicion = (str(e.idListaPrecios.codPuesto) if e.idListaPrecios.codPuesto else '') + '-' + str(e.Posicion)
                    detalle.append(puesto_posicion)

                    fecha = s.fechaInicio.strftime('%Y-%m-%d')
                    clave = (
                        str(e.Codempleado)[:8].zfill(10),
                        puesto_posicion,
                        fecha
                    )
                    print (clave)
                    if clave not in result_dict:
                        results.append(detalle)

        #creacion del archivo temporal            
        df = pd.DataFrame(results)
        file_name='turnos_javelin_%s.csv' % datetime.date.today().strftime('%Y-%m-%d')
        file_path = '%stmp/%s' % (MEDIA_ROOT,file_name)
        file_url = '%stmp/%s' % (MEDIA_URL,file_name)    
        df.to_csv(file_path, index=False, header=False)
        data =  {'file_url': file_url}
        return  JsonResponse(data)  
    else:
        return HttpResponseRedirect(LOGIN_URL)   

def inconsistencias_turnosJavelin(request):
    if request.user.is_authenticated and request.user.has_perm('solicitudes.read_solicitud'):
        detalle ={}
        results = []
        today = datetime.date.today()
        cliente = ''
        sucursal = ''
        desde = today
        hasta= today
        tipo_solicitud = ''
        estado =''
        solicitudes = []
        import re
        ##obtener los clientes
        connection = pyhdb.connect(
                host=P_SERVER,
                port=P_PORT,
                user=P_USER_SERVER,
                password=P_PASSWD_SERVER
            )
        cursor = connection.cursor()
        sql = """SELECT "CardCode","CardName" FROM "%s"."OCRD" T0 WHERE T0."CardType" in ('C') and T0."validFor"='Y' """ % (P_COMPANY)
        cursor.execute(sql)
        listado_clientes = cursor.fetchall()
        dict_clientes = {}
        for p in listado_clientes:
            dict_clientes['%s' % p[0]] = p[1]

        sql = """SELECT "CODEMPLEADO","CARGO" FROM "%s"."MAESTROS_EMPLEADO" """ % (ESQUEMA_VENTAS)
        cursor.execute(sql)
        listado_empleados = cursor.fetchall()
        dict_empleados = {}
        for p in listado_empleados:
            dict_empleados['%s' % p[0]] = p[1]
        try:
            connection.close()
        except:
            pass

        if request.method == 'POST':
            tipo_solicitud = 'SSEE'
            cliente = request.POST['cliente']
            sucursal = request.POST['sucursal']    
            desde = datetime.datetime.strptime(request.POST['desde'],'%d/%m/%Y')
            hasta = datetime.datetime.strptime(request.POST['hasta'],'%d/%m/%Y')
            fecha_desde_solo = desde.strftime('%Y-%m-%d')
            fecha_hasta_solo = hasta.strftime('%Y-%m-%d')
            estado = '3'
            
            ### VALIDACION ###
            connection = MySQLdb.connect(host=P_SERVER_JC,user=P_USER_JC,passwd=P_CLAVE_JC, port=int(P_PORT_JC))
            mysql_cursor = connection.cursor()
            sqlval = """
            select distinct
            emp.legacy_employee_identifier AS id_employee,
            sch.quick_shift_desc as cod,
            date(convert_tz(sch.shift_start_date_time,'Etc/UTC','America/Lima')) as initial_date
            from  g4s_scheduler.schedule_history sch
            left join employees_and_assets.employee emp on emp.id = sch.employee_id
            left join customers_and_contracts.WORKORDER_LOCATION wol ON wol.id = sch.work_order_location_id
            left join customers_and_contracts.WORK_ORDER wor ON wor.id = wol.work_order_id
            left join customers_and_contracts.contract ctr ON ctr.id = wor.contract_id
            left join customers_and_contracts.account acc ON acc.id = ctr.account_id
            left join customers_and_contracts.customer cus ON cus.id = acc.customer_id
            left join customers_and_contracts.billing_information bil on wor.billing_information_id=bil.id
            left join customers_and_contracts.shift_details sh
            on sh.id=sch.shift_id and sch.post_id=sh.post_id
            where acc.customerReference1 is not null 
            and bil.work_order_type_id=62
            and date(convert_tz(sch.shift_start_date_time,'Etc/UTC','America/Lima'))
            between '%s' and '%s'
            """% (
            fecha_desde_solo,
            fecha_hasta_solo)
            mysql_cursor.execute(sqlval)
            result_dict = {}

            for row in mysql_cursor.fetchall():
                id_employee, cod, initial_date,  = row
                formatted_date = initial_date.strftime('%d/%m/%Y')
                key = (id_employee, cod, formatted_date)
                result_dict[key] = {
                    'id_employee': id_employee,
                    'cod': cod,
                    'initial_date': formatted_date,
                }

            mysql_cursor.close()
            connection.close()
            print(result_dict)

            if desde != '' and hasta != '':
                    solicitudes = MVC_Solicitud_Servicio.objects.filter(fechaInicio__gte=desde,fechaFin__lte=hasta)
            if cliente != '':
                solicitudes = solicitudes.filter(idSolicitud__codCliente=cliente)
            if sucursal != '':
                qs = solicitudes.filter(codBranch = None)
                solicitudes = solicitudes.filter(codBranch_id=sucursal)
            if estado != '':
                estado = MVC_Lv_Estado_Solicitud_Servicio.objects.filter(idEstado = estado)
                if estado.count() > 0:
                    solicitudes = solicitudes.filter(idEstado=estado[0].idEstado)
        else:
            solicitudes = MVC_Solicitud_Servicio.objects.filter(fechaInicio__gte=today,fechaFin__lte=today)
    
        regis_puesto = {}
        for s in solicitudes:
            for e in s.g4s_estado_solicitud_set.all():
                if e.Soline in regis_puesto.keys():
                    pass
                else:
                    if e.Soline:
                        posDet = MVC_Os_Posicion_Det.objects.filter(idOSPosicionDet = int(e.Soline))
                        if posDet.count() > 0:
                            if posDet[0].idListaPrecios:
                                regis_puesto['%s' % e.Soline] = {'codigosaturn1':'','codigosaturn2':''}
                                regis_puesto['%s' % e.Soline]['codigoSaturn1'] = posDet[0].codigoSaturn1
                                regis_puesto['%s' % e.Soline]['codigoSaturn2'] = posDet[0].codigoSaturn2
        #emcabezado
        results.append(['solicitud','cliente','nombre_cliente','branch','site','cantidad','fecha inicio','fecha fin','hora inicio','hora fin','DNI','empleado','fecha inicio','hora incio','fecha fin','hora fin','codigo posicion'])
        #llenado de data
        if tipo_solicitud == '' or tipo_solicitud == T_SERVICIO:
            for s in solicitudes:
                for e in s.g4s_estado_solicitud_set.all():
                    # Construir la clave para buscar en result_dict
                    codEmpleado = str(e.Codempleado)[:8].zfill(10)
                    puesto_posicion = str(e.idListaPrecios.codPuesto) + '-' + str(e.Posicion)
                    fecha_inicio = s.fechaInicio.strftime('%Y-%m-%d')
                    clave = (codEmpleado, puesto_posicion, fecha_inicio)
                    print(clave)

                    if clave not in result_dict:
                        detalle = []
                        detalle.append(s.codsolicitud())
                        detalle.append(s.idSolicitud.codCliente)
                        try:
                            detalle.append(dict_clientes['%s' % s.idSolicitud.codCliente])
                        except:
                            detalle.append('')
                        detalle.append(s.branch())
                        detalle.append(s.site())
                        detalle.append(str(s.cantidad))
                        try:
                            detalle.append(s.fechaInicio.strftime('%d/%m/%Y'))
                        except:
                            detalle.append('')
                        try:
                            detalle.append(s.fechaFin.strftime('%d/%m/%Y'))
                        except:
                            detalle.append('')
                        detalle.append(s.horaInicio)
                        detalle.append(s.horaFin)
                        ######################################
                        detalle.append(str(e.Codempleado)[:9])
                        detalle.append(e.empleado())
                        try:
                            detalle.append(e.FechaInicio.strftime('%d/%m/%Y'))
                        except:
                            detalle.append('')
                        try:
                            detalle.append(e.FechaInicio.strftime('%H:%M'))
                        except:
                            detalle.append('')
                        try:
                            detalle.append(e.FechaFin.strftime('%d/%m/%Y'))
                        except:
                            detalle.append('')
                        try:
                            detalle.append(e.FechaFin.strftime('%H:%M'))
                        except:
                            detalle.append('')
                        puesto_posicion = (str(e.idListaPrecios.codPuesto) if e.idListaPrecios.codPuesto else '') + '-' + str(e.Posicion)
                        detalle.append(puesto_posicion)

                        results.append(detalle)
            
        df = pd.DataFrame(results)
        file_name='inconsistencias_turnos_javelin_%s.csv' % datetime.date.today().strftime('%Y-%m-%d')
        file_path = '%stmp/%s' % (MEDIA_ROOT,file_name)
        file_url = '%stmp/%s' % (MEDIA_URL,file_name)    
        df.to_csv(file_path, index=False, header=False)
        data =  {'file_url': file_url}
        return  JsonResponse(data)  
    else:
        return HttpResponseRedirect(LOGIN_URL)   



def solicitudes_buscar(request):
    if request.user.is_authenticated and request.user.has_perm('solicitudes.read_solicitud'):
        detalle = {}
        results = []
        today = datetime.date.today()
        codsolicitud = ''
        cliente = ''
        solicitante = ''
        desde = today
        hasta= today
        estado =''
        solicitudes = []
        import re
        patron = re.compile('SE\d{6,}')
        if request.method == 'POST':
            codsolicitud = request.POST['codsolicitud']
            try:
                codsol_s = int(codsolicitud)
                valid_codsol = True
            except:
                codsol_s = 0
                valid_codsol = False

            cliente = request.POST['cliente']
            solicitante = request.POST['solicitante']
            desde = datetime.datetime.strptime(request.POST['desde'],'%d/%m/%Y')
            hasta = datetime.datetime.strptime(request.POST['hasta'],'%d/%m/%Y')
            estado = request.POST['estado']

            if desde != '' and hasta != '':
                solicitudes = MVC_Solicitud.objects.filter(fechaEmision__gte=desde,fechaEmision__lte=hasta)
            if codsolicitud != '' and valid_codsol==True:
                solicitudes = solicitudes.filter(id=codsol_s)
            if cliente != '':
                solicitudes = solicitudes.filter(codCliente=cliente)
            if solicitante != '':
                solicitudes = solicitudes.filter(Q(idSolicitante__Apepaterno__icontains=solicitante) | Q(idSolicitante__Apematerno__icontains=solicitante) | Q(idSolicitante__Nombres__icontains=solicitante))
            if estado != '':
                estado = MVC_Lv_Estado_Solicitud_Servicio.objects.filter(idEstado = estado) #Lista esta mal debe ser MVC_Lv_Estado_Solicitud Jie 05/09/2020
                if estado.count() > 0:
                    solicitudes = solicitudes.filter(idEstado=estado[0].idEstado)
        else:
            solicitudes = MVC_Solicitud.objects.filter(fechaEmision=today)

        for s in solicitudes:
            detalle_json = {}
            detalle_json['codigo'] = s.codsolicitud()
            detalle_json['fecha'] = s.fechaEmision.strftime('%d/%m/%Y')
            detalle_json['cliente'] = s.codCliente
            detalle_json['solicitante'] = s.idSolicitante.__str__()
            detalle_json['estado'] = s.idEstado.nombre
            #detalle_json['detalle'] = s.Observaciones
            detalle_json['detalle'] = ''
            detalle_json['flujo'] = 'flujo'
            results.append(detalle_json)
        detalle['rows'] = results 
        data = json.dumps(detalle)
        mimetype = 'application/json'
        return HttpResponse(data, mimetype)
    else:
        return HttpResponseRedirect(LOGIN_URL)
    
def buscar_grm_listado_detallado_query(p_tipocot,p_tipocalculo, p_cliente,p_fechaDesde,p_fechaHasta,p_columna,p_orden,p_start,p_length):

    id_fecha=''
    texto_cliente=''
    id_tipocot='AND Q.TIPOCOTIZACION IN ('
    id_tipocalculo=''
    id_columna = str(int(p_columna) + 1)
    id_orden= p_orden
    id_start = p_start
    id_length = p_length
    id_dias = '30'


    if p_fechaDesde=='null' and p_fechaHasta=='null': #los dos estan vacio
        return ''


    elif p_fechaDesde!='null' and p_fechaHasta!='null': #los dos estan
        id_fecha='''AND (
                      ((DAYS_BETWEEN(Q.FECHAFIN, '%s')<=0) AND (DAYS_BETWEEN(Q.FECHAINICIO,'%s')>=0))
                      OR
                      ((Q.FECHAFIN IS NULL)AND(DAYS_BETWEEN(Q.FECHAINICIO,'%s')>=0)))'''%(p_fechaDesde,p_fechaHasta,p_fechaHasta)
        id_dias='''(CASE WHEN (DAYS_BETWEEN('%s','%s') <= 28 AND MONTH('%s')=2 AND MONTH('%s')=2 ) THEN 28 ELSE 30 END)  ''' % (p_fechaDesde,p_fechaHasta,p_fechaDesde,p_fechaHasta)

    elif p_fechaDesde!='null' :#FECHA DESDE
        print('--------------------------------------------')
        id_fecha='''AND (DAYS_BETWEEN(Q.FECHAFIN, '%s')<=0
                    ---Siguen activos, no estan cerrados
                     OR
                    ( Q.FECHAFIN IS NULL))'''%(p_fechaDesde)

    else:#FECHA HASTA
        id_fecha='''AND (DAYS_BETWEEN(Q.FECHAINICIO,'%s') >=0)'''%(p_fechaHasta)



    if p_cliente!='null':
        texto_cliente='''AND ((SELECT
                      UPPER(RTRIM(T9."CardName"))
                      FROM "%s"."OCRD" T9
                      WHERE T9."CardType" in ('C',
                      'L')
                      AND UPPER(T9."CardCode") =UPPER(Q.CODCLIENTE)) like '%s')'''%(P_COMPANY,p_cliente.upper())
    else:
        texto_cliente=''

    if p_tipocot!='null':
        if len(p_tipocot) != 0:
            for i in range(len(p_tipocot)):
                if i!= (len(p_tipocot) - 1):
                    if i%2 == 0:
                        id_tipocot += ''' (SELECT T6.NOMBRE FROM "%s"."VENTAS_MVC_LV_TIPO_COTIZACION" T6 WHERE T6.IDTIPOCOTIZACION='%s'),  ''' % (ESQUEMA_VENTAS,p_tipocot[i])
                    else:
                        id_tipocot += ''' (SELECT T6.NOMBRE FROM "%s"."VENTAS_MVC_LV_TIPO_COTIZACION" T6 WHERE T6.IDTIPOCOTIZACION='%s'),  ''' % (ESQUEMA_VENTAS,p_tipocot[i])
                else:
                    id_tipocot += ''' (SELECT T6.NOMBRE FROM "%s"."VENTAS_MVC_LV_TIPO_COTIZACION" T6 WHERE T6.IDTIPOCOTIZACION='%s')  ''' % (ESQUEMA_VENTAS,p_tipocot[i])

            id_tipocot += ' )'
    else:
        id_tipocot=''

    if p_tipocalculo!='null':
        if p_tipocalculo == '0':
            id_tipocalculo=''' DAYS_BETWEEN(CASE WHEN Q.FECHAINICIO < '%s' THEN '%s' ELSE Q.FECHAINICIO END,CASE WHEN Q.FECHAFIN IS NULL OR Q.FECHAFIN > '%s'  THEN '%s' ELSE Q.FECHAFIN END) + 1 ''' % (p_fechaDesde,p_fechaDesde,p_fechaHasta,p_fechaHasta)
        elif p_tipocalculo == '1':
            id_tipocalculo=id_dias
        elif p_tipocalculo == '2':
            id_tipocalculo=''' CASE WHEN Q.TIPOCOTIZACION IN ('G4S','G4S-Cliente') THEN DAYS_BETWEEN(CASE WHEN Q.FECHAINICIO < '%s' THEN '%s' ELSE Q.FECHAINICIO END,CASE WHEN Q.FECHAFIN IS NULL OR Q.FECHAFIN > '%s'  THEN '%s' ELSE Q.FECHAFIN END) + 1 ELSE (CASE WHEN %s=28 THEN 28 ELSE 30 END) END ''' % (p_fechaDesde,p_fechaDesde,p_fechaHasta,p_fechaHasta,id_dias)
    else:
        id_tipocalculo=''' DAYS_BETWEEN(CASE WHEN Q.FECHAINICIO < '%s' THEN '%s' ELSE Q.FECHAINICIO END,CASE WHEN Q.FECHAFIN IS NULL OR Q.FECHAFIN > '%s'  THEN '%s' ELSE Q.FECHAFIN END) + 1 ''' % (p_fechaDesde,p_fechaDesde,p_fechaHasta,p_fechaHasta)


    connection = pyhdb.connect(host=P_SERVER, port=P_PORT,user=P_USER_SERVER, password=P_PASSWD_SERVER)
    cursor = connection.cursor()

    ###OBTENEMOS EL TOTAL DE REGISTROS -------------------
    query_c='''
        SELECT
         COUNT(*)
        FROM "%s".REPORTE_GROSSMARGIN_CLIENTE_DATA_V AS Q
        WHERE Q.CODCLIENTE IS NOT NULL %s
                      AND (Q.FECHAINICIO<=Q.FECHAFIN OR Q.FECHAFIN IS NULL)
                      AND (Q.ESTADO in ('Confirmado','Versión Anterior'))
                      AND Q.FACTURABLE='SI'
                    %s %s ORDER BY %s %s'''%(ESQUEMA_VENTAS, id_tipocot ,id_fecha, texto_cliente, id_columna,id_orden)

    query_puecomp_c='''
        SELECT
         COUNT(*)
        FROM "%s".REPORTE_GROSSMARGIN_CLIENTE_DATA_V AS Q
        WHERE Q.CODCLIENTE IS NOT NULL %s
                      AND (Q.FECHAINICIO<=Q.FECHAFIN OR Q.FECHAFIN IS NULL)
                      AND (Q.ESTADO in ('Confirmado','Versión Anterior'))
                      AND Q.FACTURABLE='NO'
                      AND Q.TIPO IN ('SSFF')
                    %s %s ORDER BY %s %s'''%(ESQUEMA_VENTAS, id_tipocot ,id_fecha, texto_cliente, id_columna,id_orden)


    print(query_c)
    print(query_puecomp_c)
    cursor.execute(query_c)
    listado_numrows = cursor.fetchall()
    total_rows=0
    for cant in listado_numrows:
        total_rows+=cant[0]

    cursor.execute(query_puecomp_c)
    listado_numrows1 = cursor.fetchall()
    for cant in listado_numrows1:
        total_rows+=cant[0]

    #####OBTENEMOS LOS REGISTROS POR PAGINACIÓN--------------
    print("ACA ESTA EL REPORTE")
    query='''
        SELECT
         Q."TIPO",
         Q."CODCLIENTE",
         Q."NOMCLIENTE",
         Q."NROORDENSERVICIO",
         Q."VERSIONOS",
         Q."ESTADO",
         Q."NROCOTIZACION",
         Q."VERSIONCOTIZACION",
         Q."TIPOCOTIZACION",
         Q."FACTURABLE",
         Q."CODASSIGNMENT",
         Q."NOMASSIGNMENT",
         Q."FECHAINICIO",
         Q."FECHAFIN",
         Q."IDITEM",
         Q."CODITEM",
         Q."NOMBREITEM",
         Q."CODPUESTO",
         Q."PUESTO",
         Q."CANTIDAD",
         Q."MONTONOMINA",
         Q."MONTOUNIFORME",
         Q."MONTOACCCOMP",
         Q."MONTOACCPUESTO",
         Q."MONTOPUESTOCOMP",
         Q."MONTODESCANSERO",
         Q."UNIFORMEDESCANSERO",
         Q."REFRIGERANTEVALOR",
         Q."GASTOSOPER",
         Q."GASTOSADMIN",
         Q."UTILIDAD",
         Q."IMPORTEMENSUAL",
         Q."COSTOMENSUAL",
         Q."COSTOLABORALMENSUAL",
        CASE WHEN Q.FECHAINICIO < '%s' THEN '%s' ELSE Q.FECHAINICIO END AS FECHAINICIOREPORTE,
        CASE WHEN Q.FECHAFIN IS NULL OR Q.FECHAFIN > '%s'  THEN '%s' ELSE Q.FECHAFIN END AS FECHAFINREPORTE,
        %s AS DIAS,
        CASE WHEN Q.TIPO IN ('SSEE','SSEEEVE','SSEEACC') THEN Q.IMPORTEMENSUAL ELSE ROUND( (Q.IMPORTEMENSUAL/30)*(%s) ,2) END AS TOTALVENTA,
        ROUND( (Q.COSTOMENSUAL/30)*(%s) ,2) AS TOTALCOSTO,
        ROUND( (Q.COSTOLABORALMENSUAL/30)*(%s) ,2) AS TOTALCOSTOLABORAL,
        Q."CODUBIGEO",
        Q."DEPARTAMENTO",
        Q."PROVINCIA",
        Q."DISTRITO",
        Q."DIMSUCSAP",
        Q."SUCURSALG4S",
        Q."RESPONSABLESUC",
        Q."REGION",
        Q."CODPOSICION",        
        Q."CODSATURN",
        Q."CODPUESTO" ||'-'|| Q1."POSICION",
        %s AS TOTAL_ROWS,
        (SELECT T2.DESCRIPCION FROM "%s"."VENTAS_MVC_PUESTO" T0
        INNER JOIN "%s"."VENTAS_MVC_COTIZACION" T1 ON T0.IDCOTIZACION_ID=T1.IDCOTIZACION
        INNER JOIN "%s"."VENTAS_MVC_COTIZACION_SERVICIO" T2 ON T2.IDCOTIZACION_ID=T1.IDCOTIZACION
        AND T0.IDSERVICIOCOT_ID = T2.IDSERVICIOCOT
        WHERE IDPUESTO=Q.IDPUESTO ) AS PUESTO_DESC
        FROM "%s".REPORTE_GROSSMARGIN_CLIENTE_DATA_V AS Q
        LEFT JOIN "%s".VENTAS_MVC_OS_POSICION_DET Q1 ON
        TO_CHAR(Q1.IDOSPOSICIONDET) = SUBSTRING(Q.IDITEM,6)
        WHERE Q.CODCLIENTE IS NOT NULL %s
                      AND (Q.FECHAINICIO<=Q.FECHAFIN OR Q.FECHAFIN IS NULL)
                      AND (Q.ESTADO in ('Confirmado','Versión Anterior'))
                      AND Q.FACTURABLE='SI'
                    %s %s ORDER BY %s %s'''%(p_fechaDesde,p_fechaDesde,p_fechaHasta,p_fechaHasta,
                    id_tipocalculo,
                    id_tipocalculo,
                    id_tipocalculo,
                    id_dias,
                    total_rows,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS, id_tipocot ,id_fecha, texto_cliente, id_columna,id_orden)

    query_puecomp='''
        SELECT
         Q."TIPO",
         Q."CODCLIENTE",
         Q."NOMCLIENTE",
         Q."NROORDENSERVICIO",
         Q."VERSIONOS",
         Q."ESTADO",
         Q."NROCOTIZACION",
         Q."VERSIONCOTIZACION",
         Q."TIPOCOTIZACION",
         Q."FACTURABLE",
         Q."CODASSIGNMENT",
         Q."NOMASSIGNMENT",
         Q."FECHAINICIO",
         Q."FECHAFIN",
         Q."IDITEM",
         Q."CODITEM",
         Q."NOMBREITEM",
         Q."CODPUESTO",
         Q."PUESTO",
         Q."CANTIDAD",
         Q."MONTONOMINA",
         Q."MONTOUNIFORME",
         Q."MONTOACCCOMP",
         Q."MONTOACCPUESTO",
         Q."MONTOPUESTOCOMP",
         Q."MONTODESCANSERO",
         Q."UNIFORMEDESCANSERO",
         Q."REFRIGERANTEVALOR",
         Q."GASTOSOPER",
         Q."GASTOSADMIN",
         Q."UTILIDAD",
         Q."IMPORTEMENSUAL",
         Q."COSTOMENSUAL",
         Q."COSTOLABORALMENSUAL",
        CASE WHEN Q.FECHAINICIO < '%s' THEN '%s' ELSE Q.FECHAINICIO END AS FECHAINICIOREPORTE,
        CASE WHEN Q.FECHAFIN IS NULL OR Q.FECHAFIN > '%s'  THEN '%s' ELSE Q.FECHAFIN END AS FECHAFINREPORTE,
        %s AS DIAS,
        CASE WHEN Q.TIPO IN ('SSEE','SSEEEVE','SSEEACC') THEN Q.IMPORTEMENSUAL ELSE ROUND( (Q.IMPORTEMENSUAL/30)*(%s) ,2) END AS TOTALVENTA,
        ROUND( (Q.COSTOMENSUAL/30)*(%s) ,2) AS TOTALCOSTO,
        ROUND( (Q.COSTOLABORALMENSUAL/30)*(%s) ,2) AS TOTALCOSTOLABORAL,
        Q."CODUBIGEO",
        Q."DEPARTAMENTO",
        Q."PROVINCIA",
        Q."DISTRITO",
        Q."DIMSUCSAP",
        Q."SUCURSALG4S",
        Q."RESPONSABLESUC",
        Q."REGION",
        Q."CODPOSICION",   
        Q."CODSATURN",
        Q."CODPUESTO" ||'-'|| Q1."POSICION",
        %s AS TOTAL_ROWS,
        (SELECT T2.DESCRIPCION FROM "%s"."VENTAS_MVC_PUESTO" T0
        INNER JOIN "%s"."VENTAS_MVC_COTIZACION" T1 ON T0.IDCOTIZACION_ID=T1.IDCOTIZACION
        INNER JOIN "%s"."VENTAS_MVC_COTIZACION_SERVICIO" T2 ON T2.IDCOTIZACION_ID=T1.IDCOTIZACION
        AND T0.IDSERVICIOCOT_ID = T2.IDSERVICIOCOT
        WHERE IDPUESTO=Q.IDPUESTO ) AS PUESTO_DESC
        FROM "%s".REPORTE_GROSSMARGIN_CLIENTE_DATA_V AS Q
        LEFT JOIN "%s".VENTAS_MVC_OS_POSICION_DET Q1 ON
        TO_CHAR(Q1.IDOSPOSICIONDET) = SUBSTRING(Q.IDITEM,6)
        WHERE Q.CODCLIENTE IS NOT NULL %s
                      AND (Q.FECHAINICIO<=Q.FECHAFIN OR Q.FECHAFIN IS NULL)
                      AND (Q.ESTADO in ('Confirmado','Versión Anterior'))
                      AND Q.FACTURABLE='NO'
                      AND Q.TIPO IN ('SSFF')
                    %s %s ORDER BY %s %s'''%(p_fechaDesde,p_fechaDesde,p_fechaHasta,p_fechaHasta,
                    id_tipocalculo,
                    id_dias,
                    id_tipocalculo,
                    id_dias,
                    total_rows,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS,
                    ESQUEMA_VENTAS, id_tipocot ,id_fecha, texto_cliente, id_columna,id_orden)

    print(query)
    print(query_puecomp)


    cursor.execute(query)
    listado = cursor.fetchall()

    cursor.execute(query_puecomp)
    listado_puecomp = cursor.fetchall()

    resultado = listado + listado_puecomp
    connection.close()
    return resultado

def reporte_tarifa_horasextras_det(request):
    if request.user.is_authenticated:
        if request.method == 'POST':
            ruc_cliente = request.POST.get('ruc_cliente', False)
            texto_cliente = request.POST.get('texto_cliente', False)
            id_fechaDesde = request.POST.get('id_fechaDesde', False)
            id_fechaHasta = request.POST.get('id_fechaHasta', False)
            id_fechaDesde = str(datetime.strptime(id_fechaDesde, '%d/%m/%Y').date())
            id_fechaHasta = str(datetime.strptime(id_fechaHasta, '%d/%m/%Y').date())
            print(ruc_cliente)
            print(texto_cliente)
            print(id_fechaDesde)
            print(id_fechaHasta)
            filtro = ""
            if texto_cliente != "Todos":
                filtro = "AND Q.CODCLIENTE = '%s'" % (ruc_cliente)

            connection = pyhdb.connect(host = P_SERVER, port = P_PORT, user = P_USER_SERVER, password = P_PASSWD_SERVER)
            cursor = connection.cursor()
            sql = """
                    SELECT 
                    Q."TIPOCOTIZACION",
                    Q."CODCLIENTE",
                    Q."NOMCLIENTE",
                    Q."NROORDENSERVICIO",
                    Q."VERSIONOS",
                    Q."CODBRANCH",
                    Q."NOMBRANCH",
                    Q."CODASSIGNMENT",
                    (SELECT ID FROM "B1H_MYSALES".MAESTROS_ASSIGNMENT WHERE CODCLIENTE_ID || '-' || CODASSIGNMENT =  Q."CODCLIENTE" || '-' || Q."CODASSIGNMENT" ) AS IDUNIDAD,
                    Q."NOMASSIGNMENT",
                    Q."CODUBIGEO",
                    Q."DEPARTAMENTO",
                    Q."PROVINCIA",
                    Q."DISTRITO",
                    Q."DIMSUCSAP",
                    Q."SUCURSALG4S",  
                    Q."RESPONSABLESUC",
                    Q."REGION",
                    Q."FECHAINICIO",
                    Q."FECHAFIN",
                    Q."IDITEM",
                    Q."CODITEM",
                    Q."NOMBREITEM",
                    Q."CODPUESTO",
                    Q.IDPUESTO,
                    Q."CODSATURN",
                    Q."CODPUESTO"||'-'||Q1."POSICION" AS CODPOSICION,
                    Q."PUESTO",

                    (SELECT T2.DESCRIPCION FROM "B1H_MYSALES"."VENTAS_MVC_PUESTO" T0
                    INNER JOIN "B1H_MYSALES"."VENTAS_MVC_COTIZACION" T1 ON T0.IDCOTIZACION_ID=T1.IDCOTIZACION
                    INNER JOIN "B1H_MYSALES"."VENTAS_MVC_COTIZACION_SERVICIO" T2 ON T2.IDCOTIZACION_ID=T1.IDCOTIZACION
                    AND T0.IDSERVICIOCOT_ID = T2.IDSERVICIOCOT
                    WHERE IDPUESTO=Q1.IDPUESTO_ID
                    ) AS PUESTO_DESC,
                    
                    Q."CANTIDAD",
                    Q."MONTONOMINA",
                    Q."MONTOUNIFORME",
                    Q."MONTOACCCOMP",
                    Q."MONTOACCPUESTO",
                    Q."MONTOPUESTOCOMP",
                    Q."MONTODESCANSERO",
                    Q."UNIFORMEDESCANSERO",
                    Q."REFRIGERANTEVALOR",
                    Q."GASTOSOPER",
                    Q."GASTOSADMIN",
                    Q."UTILIDAD",  
                    Q."IMPORTEMENSUAL",
                    MIN(CO.TARIFAHORANOCTURNA) AS "TARIFAHHEE",
                    Q."ESTADO"

                    
                    FROM "B1H_MYSALES".REPORTE_GROSSMARGIN_CLIENTE_DATA_V AS Q
                    LEFT JOIN "B1H_MYSALES".VENTAS_MVC_OS_POSICION_DET Q1 ON 
                    TO_CHAR(Q1.IDOSPOSICIONDET) = SUBSTRING(Q.IDITEM,6)
                    LEFT JOIN "B1H_MYSALES".VENTAS_MVC_COTIZACION_OTRO CO ON CO.IDPUESTO_ID = Q.IDPUESTO

                    WHERE 
                    Q.CODCLIENTE IS NOT NULL  
                    AND (Q.FECHAINICIO<=Q.FECHAFIN OR Q.FECHAFIN IS NULL)
                    AND (Q.ESTADO in ('Confirmado','Versión Anterior'))
                    AND Q.TIPO IN ('SSFF')
                    AND (
                        ((DAYS_BETWEEN(Q.FECHAFIN, '%s')<=0) AND (DAYS_BETWEEN(Q.FECHAINICIO,'%s')>=0))
                        OR
                        ((Q.FECHAFIN IS NULL)AND(DAYS_BETWEEN(Q.FECHAINICIO,'%s')>=0))) 
                    %s
                    --and substr(Q.CODPUESTO,1,11) = '00152423961'
                    GROUP BY Q."IDITEM",Q.TIPOCOTIZACION,Q.CODCLIENTE,Q.NOMCLIENTE,Q.NROORDENSERVICIO,Q.VERSIONOS,Q.CODBRANCH,
                    Q."NOMBRANCH",
                    Q."CODASSIGNMENT",
                    Q.NOMASSIGNMENT,
                    Q.CODUBIGEO,
                    Q.DEPARTAMENTO,
                    Q.PROVINCIA,
                    Q.DISTRITO,
                    Q.DIMSUCSAP,
                    Q.SUCURSALG4S,
                    Q.RESPONSABLESUC,
                    Q.REGION,
                    Q.FECHAINICIO,
                    Q.FECHAFIN,
                    Q.CODITEM,
                    Q."NOMBREITEM",
                    Q."CODPUESTO",
                    Q.IDPUESTO,
                    Q."CODSATURN",
                    Q1."POSICION",
                    Q."PUESTO",
                    Q1.IDPUESTO_ID,
                    Q."CANTIDAD",
                    Q."MONTONOMINA",
                    Q."MONTOUNIFORME",
                    Q."MONTOACCCOMP",
                    Q."MONTOACCPUESTO",
                    Q."MONTOPUESTOCOMP",
                    Q."MONTODESCANSERO",
                    Q."UNIFORMEDESCANSERO",
                    Q."REFRIGERANTEVALOR",
                    Q."GASTOSOPER",
                    Q."GASTOSADMIN",
                    Q."UTILIDAD",  
                    Q."IMPORTEMENSUAL",
                    Q.ESTADO
                  """ % (
                id_fechaHasta,
                id_fechaDesde,
                id_fechaDesde,
                filtro,

                #AND Q.CODCLIENTE = '%s'
            )
            print(sql)
            cursor.execute(sql)
            listado = cursor.fetchall()
            print(listado)
            lista_mysql = [[
                'TIPOCOTIZACION',
                    'CODCLIENTE',
                    'NOMCLIENTE',
                    'NROORDENSERVICIO',
                    'VERSIONOS',
                    'CODBRANCH',
                    'NOMBRANCH',
                    'CODASSIGNMENT',
                    'IDUNIDAD',
                    'NOMASSIGNMENT',
                    'CODUBIGEO',
                    'DEPARTAMENTO',
                    'PROVINCIA',
                    'DISTRITO',
                    'DIMSUCSAP',
                    'SUCURSALG4S',
                    'RESPONSABLESUC',
                    'REGION',
                    'FECHAINICIO',
                    'FECHAFIN',
                    'IDITEM',
                    'CODITEM',
                    'NOMBREITEM',
                    'CODPUESTO',
                    'IDPUESTO',
                    'CODSATURN',
                    'CODPOSICION',
                    'PUESTO',
                    'PUESTO_DESC',
                    'CANTIDAD',
                    'MONTONOMINA',
                    'MONTOUNIFORME',
                    'MONTOACCCOMP',
                    'MONTOACCPUESTO',
                    'MONTOPUESTOCOMP',
                    'MONTODESCANSERO',
                    'UNIFORMEDESCANSERO',
                    'REFRIGERANTEVALOR',
                    'GASTOSOPER',
                    'GASTOSADMIN',
                    'UTILIDAD',
                    'IMPORTEMENSUAL',
                    'TARIFAHHEE',
                    'ESTADO',
            ]]
            for r in listado:
                lista_mysql.append(r)
            encabezado = lista_mysql[0]
            connection.close()

            df_main = pd.DataFrame(lista_mysql[1:], columns = encabezado)

            # Convertir la columna 'fechainicio' a tipo datetime
            df_main['FECHAINICIO'] = pd.to_datetime(df_main['FECHAINICIO'])
            df_main['FECHAFIN'] = pd.to_datetime(df_main['FECHAFIN'])
            # Extraer solo la parte de la fecha
            df_main['FECHAINICIO'] = df_main['FECHAINICIO'].dt.date
            df_main['FECHAFIN'] = df_main['FECHAFIN'].dt.date
            
            df_main['CANTIDAD'] = df_main['CANTIDAD'].astype(int)
            df_main['MONTONOMINA'] = df_main['MONTONOMINA'].astype(float)
            df_main['MONTOUNIFORME'] = df_main['MONTOUNIFORME'].astype(float)
            df_main['MONTOACCCOMP'] = df_main['MONTOACCCOMP'].astype(float)
            df_main['MONTOACCPUESTO'] = df_main['MONTOACCPUESTO'].astype(float)
            df_main['MONTOPUESTOCOMP'] = df_main['MONTOPUESTOCOMP'].astype(float)
            df_main['MONTODESCANSERO'] = df_main['MONTODESCANSERO'].astype(float)
            df_main['UNIFORMEDESCANSERO'] = df_main['UNIFORMEDESCANSERO'].astype(float)
            df_main['REFRIGERANTEVALOR'] = df_main['REFRIGERANTEVALOR'].astype(float)
            df_main['GASTOSOPER'] = df_main['GASTOSOPER'].astype(float)
            df_main['GASTOSADMIN'] = df_main['GASTOSADMIN'].astype(float)
            df_main['UTILIDAD'] = df_main['UTILIDAD'].astype(float)
            df_main['IMPORTEMENSUAL'] = df_main['IMPORTEMENSUAL'].astype(float)
            df_main['TARIFAHHEE'] = df_main['TARIFAHHEE'].astype(float)
            
            #codigo para saber si es tipo numerico
            #tipo_de_dato = df_main['MONTONOMINA'].dtype
            #es_numerico = pd.api.types.is_numeric_dtype(df_main['MONTONOMINA'])
            #print("Tipo de dato:", tipo_de_dato)
            #print("¿Es numérico?", es_numerico)
            print(df_main)

            #guardando el archivo en la temporal
            file_name = 'reporte_horas_extras%s.xlsx' % datetime.now().strftime('%Y-%m-%d %H%M%S')
            file_path = '%stmp/%s' % (MEDIA_ROOT, file_name)
            file_url = '%stmp/%s' % (MEDIA_URL, file_name)
            df_main.to_excel(file_path, index = False)

            data = { 'file_url': file_url }
            return JsonResponse(data)
    else:
        return HttpResponseRedirect(LOGIN_URL)