Imports System.Data
Imports System.Data.SqlClient
Module Variables

    Public TOTAL_SALARIOS As Double
    Public I As Integer

    Public NOMBRE_DEL_REPORTE_CR As String
    Public NOMBRE_DEL_REPORTE_EX As String
    Public FILTRO_SELECCION_EN_REPORTES As String

    Public AUTOMATICO As Boolean
    'VARIABLE QUE ALMACENA FECHA DE SERVIDOR
    Public FECHA_SERVIDOR As System.DateTime
    'VARIABLE QUE ALMACENA HORA DE SERVIDOR
    Public HORA_SERVIDOR As System.DateTime
    Public HORA As String
    'VARIABLES DE CONTROL DE ACTIVACION DE VENTANAS PRINCIPALES
    Public USUARIOS As Boolean      '*
    Public ESTRUCTURA As Boolean    '*
    Public MOVIMIENTOS As Boolean   '*
    Public ASPIRANTES As Boolean    '*
    Public EXPEDIENTES As Boolean   '*
    Public EXPEDIENTES_I As Boolean '*
    Public ACREDITADOS As Boolean   '*
    Public BECADOS As Boolean       '*
    Public INFORMES As Boolean      '*
    Public CONSULTA_GENERAL As Boolean       '*
    Public CONSULTA_SAL_VAC As Boolean       '*
    Public CONSULTA_MATRIZ_BAC As Boolean    '*
    Public CONSULTA_PERSONAL_AUSENTE As Boolean    '*
    Public COBERTURA_AFILIADOS As Boolean    '*
    Public TIEMPO_EXTRAORDINARIO As Boolean    '*
    Public SITUACIONES As Boolean   '*
    Public DPC1 As Boolean          '*
    Public DPC2 As Boolean          '*
    Public CATALOGOS As Boolean     '*
    Public ESTADISTICAS As Boolean  '*
    Public CONSULTAS As Boolean     '*
    Public HIGIENE_1 As Boolean     '*
    Public HIGIENE_2 As Boolean     '*
    Public HIGIENE_3 As Boolean     '*
    Public TERCERIZADOS As Boolean  '*
    Public ROLES As Boolean         '*
    '-----
    'ADAPTADOR PUBLICO
    Public MiDataAdapter As SqlDataAdapter
    'CONJUNTO DE DATOS EN MEMORIA
    Public MiDataSet As DataSet
    'CONTROL CIERRE DE VENTANAS
    Public CERRAR As Boolean
    Public CERRAR_1 As Boolean
    'RESPUESTA A MENSAJE
    Public MENSAJE
    Public MENSAJE1
    'VARIABLE DE CODIGO DE MODULO
    Public xMODULO As String
    'VARIABLE TIPO DE PERMISO
    Public PARAMETRO As Integer
    'VARIABLE QUE ALMACENA SELECCION REPORTES
    Public SELECCION As String
    Public DiccionarioParametros As New Dictionary(Of String, String)
    'VARIABLE QUE ALMACENA SELECCION DE PARAMETROS PARA REPORTES
    Public SELECCION_PARAMETRO As String
    Public USUARIO_IMPRIME As String
    'SELECCIONES DE DATOS
    Public CADENAsql As String
    Public CADENAsql1 As String
    Public CADENAsql2 As String
    Public CADENAsql3 As String
    Public CADENAsql4 As String
    Public CADENAsql5 As String
    'SELECCIONES DE NIVELES
    Public CADENA_NIVELES As String
    Public NIVEL1sql As String
    Public NIVEL2sql As String
    Public NIVEL3sql As String
    Public NIVEL4sql As String
    Public NIVEL5sql As String
    Public NIVEL6sql As String
    Public NIVEL7sql As String
    Public NIVEL8sql As String
    'VARIABLE DE ACCESO A MODULOS
    Public HAY_ACCESO As Boolean
    'VARIABLES DE NODOS
    Public NODO_TAG As Integer
    Public NODO_TEXT As String
    'VARIABLES PARAMETROS PARA REPORTES
    Public VALOR01 As String
    Public VALOR02 As String
    Public VALOR03 As String
    Public VALOR04 As String
    Public VALOR05 As String
    Public VALOR06 As String
    Public VALOR07 As String
    Public VALOR08 As String
    Public VALOR09 As String
    Public VALOR10 As String
    Public VALOR11 As String
    Public VALOR12 As String
    Public VALOR13 As String
    Public VALOR14 As String
    Public VALOR15 As String
    Public VALOR16 As String
    Public VALOR17 As String

    'Imagen dinamica
    Public VALOR18 As Boolean
    Public VALOR19 As Boolean

    Public vIP As String

    Public CADENAsqlCombo As String

    Public DATO_EN_RRHH As Boolean

    Public bPREFIJO As String
    'VARIABLES DE BUSQUEDA DE EMPLEADOS EN SPYC
    Public bEMPLEADO As String
    Public bCODIGO As String
    Public bEMPLEADO_A As String
    Public bCARGO As String
    Public bD_N1 As String
    Public bD_N2 As String
    Public bD_N3 As String
    Public bD_N4 As String
    Public bD_N5 As String
    Public bD_N6 As String
    Public bD_N7 As String
    Public bD_N8 As String
    Public bCUENTA As String
    Public cID_M_C As String
    Public cEMPLEADO As String
    Public cFECHAING As String
    Public bAN As String
    Public cORDENPAME As String
    Public CODIGO As Integer
    Public vID_M_P As Integer
    Public vID_M_C As Integer
    Public cID_HIST_M As Integer
    Public cD_T_M_EXP As String

    Public FECHA1, FECHA2 As String
    Public FECHA01, FECHA02 As String

    Public cN1 As String
    Public cN2 As String
    Public cN3 As String
    Public cN4 As String
    Public cN5 As String
    Public cN6 As String
    Public cN7 As String
    Public cN8 As String

    Public D2 As String
    Public M2 As String
    Public A2 As String
    Public D1 As String
    Public M1 As String
    Public A1 As String

    Public CE As String
    Public emp As String
    Public gr As String
    Public TIPO_INFORME As String

    Public id_EMPLEADO As Integer
    Public D, S, G As String
    Public DATO_1 As Boolean
    Public vINACTIVO As Boolean

    Public vDET1, vDET2, vDET3 As String

    Public NIVEL1x, NIVEL2x As String

    Public HAY_CARGOS As Boolean

    Public exp_PARENTELA_DIRECCION As String

    Public CODIGO_REPORTE As String
    Public NOMBRE_REPORTE As String

End Module
