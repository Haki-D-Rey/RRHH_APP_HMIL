Module Variables_Catalogos
    'CATALOGOS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'CAT DE TIPO DE ESTRUCTURA
    Public c1ID_T_ES As Integer
    Public c1DESCRIPCION As String
    Public c1SIGLAS As String
    'VISTA CAT DE ESTRUCTURA NIVEL 1
    Public c2ID_N1 As Integer
    Public c2ORDEN As String
    Public c2DESCRIPCION As String
    Public c2SIGLAS As String
    Public c2ID_T_ES As Integer
    Public c2D_T_ES As String
    Public c2MIXTA As Boolean
    Public c2MILITAR As Boolean
    Public c2PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 2
    Public c3ID_N2 As Integer
    Public c3ID_N1 As Integer
    Public c3ORDN1 As String
    Public c3D_N1 As String
    Public c3MIXTAN1 As Boolean
    Public c3MILITARN1 As Boolean
    Public c3PAMEN1 As Boolean
    Public c3ORDN2 As String
    Public c3DESCRIPCION As String
    Public c3SIGLAS As String
    Public c3ID_T_ES As Integer
    Public c3D_T_ES As String
    Public c3MIXTA As Boolean
    Public c3MILITAR As Boolean
    Public c3PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 3
    Public n3ID_N1 As Integer
    Public n3ORDN1 As String
    Public n3D_N1 As String
    Public n3ID_N2 As Integer
    Public n3ORDN2 As String
    Public n3D_N2 As String
    Public n3ID_N3 As Integer
    Public n3ORDN3 As String
    Public n3D_N3 As String
    Public n3SIGLAS As String
    Public n3ID_T_ES As Integer
    Public n3D_T_ES As String
    Public n3MIXTA As Boolean
    Public n3MILITAR As Boolean
    Public n3PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 4
    Public n4ID_N1 As Integer
    Public n4ORD1 As String
    Public n4D_N1 As String
    Public n4ID_N2 As Integer
    Public n4ORD2 As String
    Public n4D_N2 As String
    Public n4ID_N3 As Integer
    Public n4ORD3 As String
    Public n4D_N3 As String
    Public n4ID_N4 As Integer
    Public n4ORD4 As String
    Public n4D_N4 As String
    Public n4SIGLAS As String
    Public n4ID_T_ES As Integer
    Public n4D_T_ES As String
    Public n4MIXTA As Boolean
    Public n4MILITAR As Boolean
    Public n4PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 5
    Public n5ID_N1 As Integer
    Public n5ORD1 As String
    Public n5D_N1 As String
    Public n5ID_N2 As Integer
    Public n5ORD2 As String
    Public n5D_N2 As String
    Public n5ID_N3 As Integer
    Public n5ORD3 As String
    Public n5D_N3 As String
    Public n5ID_N4 As Integer
    Public n5ORD4 As String
    Public n5D_N4 As String
    Public n5ID_N5 As Integer
    Public n5ORD5 As String
    Public n5D_N5 As String
    Public n5SIGLAS As String
    Public n5ID_T_ES As Integer
    Public n5D_T_ES As String
    Public n5MIXTA As Boolean
    Public n5MILITAR As Boolean
    Public n5PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 6
    Public n6ID_N1 As Integer
    Public n6ORD1 As String
    Public n6D_N1 As String
    Public n6ID_N2 As Integer
    Public n6ORD2 As String
    Public n6D_N2 As String
    Public n6ID_N3 As Integer
    Public n6ORD3 As String
    Public n6D_N3 As String
    Public n6ID_N4 As Integer
    Public n6ORD4 As String
    Public n6D_N4 As String
    Public n6ID_N5 As Integer
    Public n6ORD5 As String
    Public n6D_N5 As String
    Public n6ID_N6 As Integer
    Public n6ORD6 As String
    Public n6D_N6 As String
    Public n6SIGLAS As String
    Public n6ID_T_ES As Integer
    Public n6D_T_ES As String
    Public n6MIXTA As Boolean
    Public n6MILITAR As Boolean
    Public n6PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 7
    Public n7ID_N1 As Integer
    Public n7ORD1 As String
    Public n7D_N1 As String
    Public n7ID_N2 As Integer
    Public n7ORD2 As String
    Public n7D_N2 As String
    Public n7ID_N3 As Integer
    Public n7ORD3 As String
    Public n7D_N3 As String
    Public n7ID_N4 As Integer
    Public n7ORD4 As String
    Public n7D_N4 As String
    Public n7ID_N5 As Integer
    Public n7ORD5 As String
    Public n7D_N5 As String
    Public n7ID_N6 As Integer
    Public n7ORD6 As String
    Public n7D_N6 As String
    Public n7ID_N7 As Integer
    Public n7ORD7 As String
    Public n7D_N7 As String
    Public n7SIGLAS As String
    Public n7ID_T_ES As Integer
    Public n7D_T_ES As String
    Public n7MIXTA As Boolean
    Public n7MILITAR As Boolean
    Public n7PAME As Boolean
    'VISTA CAT DE ESTRUCTURA NIVEL 8
    Public n8ID_N1 As Integer
    Public n8ORD1 As String
    Public n8D_N1 As String
    Public n8ID_N2 As Integer
    Public n8ORD2 As String
    Public n8D_N2 As String
    Public n8ID_N3 As Integer
    Public n8ORD3 As String
    Public n8D_N3 As String
    Public n8ID_N4 As Integer
    Public n8ORD4 As String
    Public n8D_N4 As String
    Public n8ID_N5 As Integer
    Public n8ORD5 As String
    Public n8D_N5 As String
    Public n8ID_N6 As Integer
    Public n8ORD6 As String
    Public n8D_N6 As String
    Public n8ID_N7 As Integer
    Public n8ORD7 As String
    Public n8D_N7 As String
    Public n8ID_N8 As Integer
    Public n8ORD8 As String
    Public n8D_N8 As String
    Public n8SIGLAS As String
    Public n8ID_T_ES As Integer
    Public n8D_T_ES As String
    Public n8MIXTA As Boolean
    Public n8MILITAR As Boolean
    Public n8PAME As Boolean
    'CAT DE CARGOS DE ESTRUCTURA
    Public c4ID_CARGO_ES As Integer
    Public c4DESCRIPCION As String
    'CAT DE ESPECIALIDADES MEDICAS
    Public c5ID_T_ESTUDIO As Integer
    Public c5D_T_ESTUDIO As String
    Public c5ID_E_MED As Integer
    Public c5DESCRIPCION As String
    'VISTA CAT DE SUB ESPECIALIDADES MEDICAS
    Public c6ID_E_MED As Integer
    Public c6D_E_MED As String
    Public c6ID_SE_MED As Integer
    Public c6DESCRIPCION As String
    'CAT DE GRADO PLANTILLA
    Public c7ID_GP As Integer
    Public c7DESCRIPCION As String
    Public c7SIGLAS As String
    'CAT DE GRADO REAL
    Public c8ID_GR As Integer
    Public c8DESCRIPCION As String
    Public c8SIGLAS As String
    'CAT DE NIVEL ACADEMICO
    Public c9ID_N_ACADEMICO As Integer
    Public c9DESCRIPCION As String
    'CAT DE NIVEL PROFESIONAL
    Public c10ID_N_PROFESIONAL As Integer
    Public c10DESCRIPCION As String
    Public c10SIGLAS As String
    'CAT DE NACIONALIDAD_N
    Public c11ID_NACIONALIDAD_N As Integer
    Public c11DESCRIPCION As String
    'CAT DE NACIONALIDAD_D
    Public c12ID_NACIONALIDAD_D As Integer
    Public c12DESCRIPCION As String
    'CAT DE DPTO_N
    Public c13ID_DPTO_N As Integer
    Public c13DESCRIPCION As String
    'CAT DE DPTO_D
    Public c14ID_DPTO_D As Integer
    Public c14DESCRIPCION As String
    'VISTA CAT DE MUNICIPIO_N
    Public c15ID_DPTO_N As Integer
    Public c15D_DPTO_N As String
    Public c15ID_M_N As Integer
    Public c15DESCRIPCION As String
    'VISTA CAT DE MUNICIPIO_D
    Public c16ID_DPTO_D As Integer
    Public c16D_DPTO_D As String
    Public c16ID_M_D As Integer
    Public c16DESCRIPCION As String
    'CAT DE ESTADO CIVIL
    Public c17ID_E_CIVIL As Integer
    Public c17DESCRIPCION As String
    'CAT DE TIPO DE HORARIO
    Public c18ID_T_HORARIO As Integer
    Public c18DESCRIPCION As String
    Public c18HORA_E As String
    Public c18HORA_S As String
    Public c18CHRT As Integer
    Public c18DT As Integer
    Public c18DL As Integer
    Public c18F_CAMBIA As Boolean
    'CAT DE TIPO DE AFILIACION
    Public c19ID_T_AFILIACION As Integer
    Public c19DESCRIPCION As String
    Public c19P_PATRONAL As String
    Public c19P_LABORAL As String
    'CAT DE TIPO DE PLAZA
    Public c20ID_T_PLAZA As Integer
    Public c20DESCRIPCION As String
    'CAT DE CATEGORIA SALARIAL
    Public c21ID_CAT_SALARIAL As Integer
    Public c21DESCRIPCION As String
    'CAT DE PARENTELA
    Public c22ID_PARENTELA As Integer
    Public c22DESCRIPCION As String
    'CAT DE TIPO DE ESTUDIO
    Public c23ID_T_ESTUDIO As Integer
    Public c23DESCRIPCION As String
    'CAT DE HABILIDADES ADMINISTRATIVAS
    Public c24ID_HABILIDAD As Integer
    Public c24DESCRIPCION As String
    'CAT DE IDIOMAS
    Public c25ID_IDIOMA As Integer
    Public c25DESCRIPCION As String
    'CAT DE NOMINA
    Public c26ID_NOMINA As Integer
    Public c26DESCRIPCION As String
    Public c26SIGLAS As String
    'CAT DE MOVIMIENTO EN NOMINA
    Public c27ID_MOVIMIENTO_N As Integer
    Public c27DESCRIPCION As String
    'CAT DE EVENTUALIDADES
    Public c28ID_E As Integer
    Public c28DESCRIPCION As String
    'CAT DE TIPO CURVA Y TALLA
    Public c29ID_T_CT As Integer
    Public c29DESCRIPCION As String
    'CAT DE CLASIFICACION DE PERSONAL
    Public c30ID_C_PE As Integer
    Public c30DESCRIPCION As String
    'CAT DE CLASIFICACION DE CONECORACIONES
    Public c31ID_C_CONDE As Integer
    Public c31DESCRIPCION As String
    'VISTA CAT DE CONDECORACIONES
    Public c32ID_C_CONDE As Integer
    Public c32D_C_CONDE As String
    Public c32ID_CONDE As Integer
    Public c32CODIGO As String
    Public c32DESCRIPCION As String
    'VISTA CAT DE SUB TIPO DE MOVIMIENTO REAL DE EXP
    Public c33ID_T_M_EXP As Integer
    Public c33D_T_M_EXP As String
    Public c33ID_ST_M_R_EXP As Integer
    Public c33D_ST_M_R_EXP As String
    'VISTA CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP
    Public c34ID_T_M_EXP As Integer
    Public c34D_T_M_EXP As String
    Public c34ID_ST_M_L_EXP As Integer
    Public c34D_ST_M_L_EXP As String
    'CAT DE CLASIFICACION DE PERSONAL
    Public c35ID_PAIS As Integer
    Public c35DESCRIPCION As String
    'VISTA CAT DE ANEXOS
    Public c36ID_T_ANEXO As Integer
    Public c36D_T_ANEXO As String
    Public c36ID_ANEXO As Integer
    Public c36D_ANEXO As String
    'CAT DE TIPO DE ANEXOS
    Public c37ID_T_ANEXO As Integer
    Public c37DESCRIPCION As String
    'CAT DE TIPO DE DOCUMENTO
    Public c38ID_T_DOC As Integer
    Public c38DESCRIPCION As String
    'CAT DE TIPO SANGUINEO
    Public c39ID_T_SANGUINEO As Integer
    Public c39DESCRIPCION As String
    'CAT DE CATEGORIA DE LIC
    Public c40ID_C_LIC As Integer
    Public c40DESCRIPCION As String
    'CAT DE IPSS
    Public c41ID_IPSS As Integer
    Public c41DESCRIPCION As String
    'VISTA MAESTRO DE DOCUMENTOS DIGITALES  (12)
    Public vID_M_DD As Integer
    'Public vID_M_P As Integer
    Public vID_D1 As Integer
    Public vDIRECTORIO As String
    Public vID_T_DOC As Integer
    Public vD_T_DOC As String
    Public vNOMBRE As String
    Public vOBSERVACION As String
    'CAT DE TIPO DE SITUACION DE PERSONAL
    Public c42ID_T_SIT_P As Integer
    Public c42DESCRIPCION As String
    'VISTA CAT DE SITUACION DE PERSONAL
    Public c43ID_F As Integer
    Public c43ID_T_SIT_P As Integer
    Public c43D_T_SIT_P As String
    Public c43ID_SIT_P As Integer
    Public c43D_SIT_P As String
    Public c43SIGLA As String
    'CAT DE FUNCIONALIDAD
    Public c44ID_F As Integer
    Public c44DESCRIPCION As String
    'CAT DE CONTRATOS
    Public c45ID_CONT As Integer
    Public c45DESCRIPCION As String
    Public c45ACTIVO As Boolean
    'CAT DE NIVEL DE COMPETENCIA
    Public c46ID_N_COMPETENCIA As Integer
    Public c46DESCRIPCION As String
    'CAT DE ESTABLECIMIENTOS
    Public c47ID_ESTABLECIMIENTO As Integer
    Public c47DESCRIPCION As String
    'CAT DE PREFIJOS
    Public c48ID_PREF As Integer
    Public c48DESCRIPCION As String
    'CAT DE VACUNAS
    Public c49ID_VACUNA As Integer
    Public c49DESCRIPCION As String
    'CAT DE DOSIS
    Public c50ID_DOSIS As Integer
    Public c50DESCRIPCION As String
    ' VISTA MAESTRO DE USUARIOS DEPARTAMENTOS SERVICIOS
    Public aID_AS As Integer
    Public aID_USUARIO As Integer
    Public aCUENTA As String
    Public aID_N1 As Integer
    Public aD_N1 As String
    Public aID_N2 As Integer
    Public aD_N2 As String
    'VISTA CAT DE DESTINATARIOS
    Public c44ID_DESTINATARIO As Integer
    Public c44ID_M_P As Integer
    Public c44EMPLEADO As String
    Public c44DEPARTAMENTO As String
    Public c44SERVICIO As String
    Public c44CARGO As String
    Public c44GRADO As String
    Public c44CORREO As String
    Public c44ENVIAR As Boolean
    'CAT DE FERIADOS
    Public cID_D_F As Integer
    Public cFECHA_F As String
    Public cACTIVO As Boolean
    'CAT DE ADMINISTRADORES
    Public c46ID_ADMIN As Integer
    Public c46FDESCRIPCION As String
    Public c46ACTIVO As Boolean

    Public maOM As String
    Public idmaOM As Integer
    Public maCARGO As String
    Public ADDUPD As Boolean
End Module
