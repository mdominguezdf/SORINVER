Attribute VB_Name = "Module1"
'Objetos para ADO

'DataGrid     Microsoft DataGrid Control 6.0 (OLEDB)
'DataList     Microsoft Data List Controls 6.0 (OLEDB)
'DataCombo    Microsoft Data List Controls 6.0 (OLEDB)
'DBList       Microsoft Data Bound List Controls 6.0
'DBCombo      Microsoft Data Bound List Controls 6.0
'MSHFlexGrid  Microsoft Hierarchical FlexGrid Control 6.0 (OLEDB)

' Declaramos la función que es necesaria para la función DownloadFile
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

' Variables para FSO "FileSystemObject" hace falta referencia a Microsoft Scripting Runtime
Public Obj_TextStream As Scripting.TextStream
Public obj_FSO As New Scripting.FileSystemObject
Public Obj_File As File

Public VarNumeroLineas As Long 'Para control avance en lectura importacion lineas cotizaciones

Public FicheroUDLSQL As String
Public VarMensajes As String

Public VarPreLinkYahoo As String
Public VarPostLinkYahoo As String
Public VarLinkYahoo As String

Private VarNombreFichero As String
Private VarNombreURL As String
Private VarTexto As String
Private VarCampo As String

Private VarTicker As String
Private VarNombreEmpresa As String

Private VarLin As Integer
Private VarCol As Integer

Public MatrizMercados()
Public VarLinMatrizMercados As Integer

Public MatrizAcciones()
Public VarLinMatrizAcciones As Integer

Public MatrizTicker()
Public VarLinMatrizTicker As Integer

Public MatrizTemporal()
Public VarLinMatrizTemporal As Integer

Public MatrizAvisos()

Public ConexionSQL As ADODB.Connection      'Para control conexion SQL
Public RegistroSQL As ADODB.Recordset       'Para control registros SQL
Public ResultadoSQL As ADODB.Recordset
Public RutinaSQL As String

Private NumCampo As Integer

Private CotFecha As String
Private CotApertura As String
Private CotMaximo As String
Private CotMinimo As String
Private CotCierre As String
Private CotVolumen As String
Private CotTicker As String
Private CotHora As String
Private CotPor As String

Public VarFechaDesde As String
Public VarFechaHasta As String

Public TotalProcedimiento As Double
Public UnitarioProcedimiento As Double

Public UnidadesProcedimiento As Long
Public UnidadesTotalesProcedimiento As Long

Public TamañoOriginalFichero As Long
Public TamañoFicheroAntes As Long
Public TamañoFicheroDespues As Long

Public UnitarioDetalleProcedimiento As Double

Public FechaSistema As String
Public FechaMercados As String
Public RespuestaMensaje As String

Public RutaRegistro As String

Public RutaAvisosMercados As String
Public RutaAvisosAcciones As String

Public FSoporte As String
Public CSoporte As String
Public FResistencia As String
Public CResistencia As String

Public PosSoporte As Integer
Public PosResistencia As Integer

Public PosSoporteUltimo As Integer
Public PosResistenciaUltima As Integer

Public DifSoporte As Double
Public DifResistencia As Double

Public ExisteRegistro As Boolean

Public RecorridoCotizacion As Double
Public CuerpoCotizacion As Double

Public MatrizVelas()

Public Lane_Periodo As Integer
Public Lane_K As Integer
Public Lane_D As Integer
Public Lane_DS As Integer
Public Lane_DSS As Integer
Public Lane_SobreCompra As Integer
Public Lane_SobreVenta As Integer
Public Lane_AvisoClasica_Lento As Boolean
Public Lane_AvisoClasica_Rapido As Boolean
Public Lane_AvisoSZona_Lento As Boolean
Public Lane_AvisoSZona_Rapido As Boolean
Public Lane_AvisoPopCorn_Lento As Boolean
Public Lane_AvisoPopCorn_Rapido As Boolean

Public RSI_Periodo As Integer
Public RSI_SobreCompra As Integer
Public RSI_SobreVenta As Integer
Public RSI_AvisoSalidaZona As Boolean
Public RSI_AvisoFailureSwing As Boolean
Public RSI_AvisoDivergencia As Boolean

Public CorreccionOrdenada_Aviso As Boolean
Public GapApertura_Aviso As Boolean
Public PausaHarami_Aviso As Boolean
Public PausaCasiHarami_Aviso As Boolean
Public Exceso_Aviso As Boolean
Public FuegoPaja_Aviso As Boolean
Public Doji_Aviso As Boolean
Public Martillo_Aviso As Boolean
Public EstrellaFugaz_Aviso As Boolean
Public Harami_Aviso As Boolean
Public Cobertura_Aviso As Boolean
Public Penetrante_Aviso As Boolean
Public Envolventes_Aviso As Boolean
Public Peonza_Aviso As Boolean
Public GapTasuki_Aviso As Boolean
Public GemelosBlancos_Aviso As Boolean
Public LineasSeparacion_Aviso As Boolean
Public LineasUnion_Aviso As Boolean
Public Puntapie_Aviso As Boolean
Public TresRios_Aviso As Boolean


Public MinimoReferencia As Double
Public MaximoReferencia As Double

Public MatrizLane()
Public MatrizRSI()

Public ControlDblClick As String

Public ValorNombre As String
Public ValorTickerYahoo As String
Public ValorZona As String
Public ValorControlHis As Boolean
Public ValorControlCot As Boolean
Public ValorControlValorHis As Boolean
Public ValorControlValorCot As Boolean

Public Sub FichaValor(Tipo As String, Titulo As String, Id As String)

VarTickerFichaValor = Id


Erase MatrizTemporal

ReDim MatrizTemporal(1, 27, 10)

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT * FROM " & Tipo & " LEFT OUTER JOIN " & Tipo & "ATecnico ON " & Tipo & "ATecnico.Id_" & Tipo & " = " & Tipo & ".Id_" & Tipo & " LEFT OUTER JOIN " & Tipo & "_Indicadores ON " & Tipo & "_Indicadores.Id_" & Tipo & " = " & Tipo & ".Id_" & Tipo & " WHERE TickerYahoo = '" & Id & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   ValorNombre = RegistroSQL("Nombre")
   ValorTickerYahoo = RegistroSQL("TickerYahoo")
   
   If Tipo = "Mercados" Then
    
      ValorZona = RegistroSQL("Zona")
      ValorControlHis = RegistroSQL("ControlHis")
      ValorControlCot = RegistroSQL("ControlCot")
      ValorControlValorHis = RegistroSQL("ControlValorHis")
      ValorControlValorCot = RegistroSQL("ControlValorCot")
      
   Else
      
      ValorZona = ""
   
   End If
   
   MatrizTemporal(1, 1, 2) = RegistroSQL("FechaDatos")
   
   MatrizTemporal(1, 2, 1) = RegistroSQL("SignoMM20")
   MatrizTemporal(1, 3, 1) = RegistroSQL("SignoMM50")
   MatrizTemporal(1, 4, 1) = RegistroSQL("SignoMM200")
   
   MatrizTemporal(1, 2, 2) = FormatNumber(RegistroSQL("MM20"), 2, True, False, True)
   MatrizTemporal(1, 3, 2) = FormatNumber(RegistroSQL("MM50"), 2, True, False, True)
   MatrizTemporal(1, 4, 2) = FormatNumber(RegistroSQL("MM200"), 2, True, False, True)
   
   For i = 1 To 4
   
       MatrizTemporal(1, 2, i + 2) = FormatNumber(RegistroSQL("MM20n" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 3, i + 2) = FormatNumber(RegistroSQL("MM50n" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 4, i + 2) = FormatNumber(RegistroSQL("MM200n" & CStr(i)), 2, True, False, True)
   
   Next

   MatrizTemporal(1, 5, 1) = FormatNumber(RegistroSQL("VolM20"), 2, True, False, True)
   MatrizTemporal(1, 5, 2) = FormatNumber(RegistroSQL("VolMin20"), 2, True, False, True)
   MatrizTemporal(1, 5, 3) = FormatNumber(RegistroSQL("VolMax20"), 2, True, False, True)

   MatrizTemporal(1, 6, 1) = FormatNumber(RegistroSQL("VolM50"), 2, True, False, True)
   MatrizTemporal(1, 6, 2) = FormatNumber(RegistroSQL("VolMin50"), 2, True, False, True)
   MatrizTemporal(1, 6, 3) = FormatNumber(RegistroSQL("VolMax50"), 2, True, False, True)

   MatrizTemporal(1, 7, 1) = FormatNumber(RegistroSQL("VolM200"), 2, True, False, True)
   MatrizTemporal(1, 7, 2) = FormatNumber(RegistroSQL("VolMin200"), 2, True, False, True)
   MatrizTemporal(1, 7, 3) = FormatNumber(RegistroSQL("VolMax200"), 2, True, False, True)

   MatrizTemporal(1, 8, 1) = FormatNumber(RegistroSQL("VelaM20"), 2, True, False, True)
   MatrizTemporal(1, 8, 2) = FormatNumber(RegistroSQL("VelaMin20"), 2, True, False, True)
   MatrizTemporal(1, 8, 3) = FormatNumber(RegistroSQL("VelaMax20"), 2, True, False, True)

   MatrizTemporal(1, 9, 1) = FormatNumber(RegistroSQL("VelaM50"), 2, True, False, True)
   MatrizTemporal(1, 9, 2) = FormatNumber(RegistroSQL("VelaMin50"), 2, True, False, True)
   MatrizTemporal(1, 9, 3) = FormatNumber(RegistroSQL("VelaMax50"), 2, True, False, True)

   MatrizTemporal(1, 10, 1) = FormatNumber(RegistroSQL("VelaM200"), 2, True, False, True)
   MatrizTemporal(1, 10, 2) = FormatNumber(RegistroSQL("VelaMin200"), 2, True, False, True)
   MatrizTemporal(1, 10, 3) = FormatNumber(RegistroSQL("VelaMax200"), 2, True, False, True)
   
   MatrizTemporal(1, 11, 1) = FormatNumber(RegistroSQL("Soporte20"), 2, True, False, True)
   MatrizTemporal(1, 11, 2) = RegistroSQL("FechaSoporte20")
   MatrizTemporal(1, 11, 3) = FormatNumber(RegistroSQL("Resistencia20"), 2, True, False, True)
   MatrizTemporal(1, 11, 4) = RegistroSQL("FechaResistencia20")

   MatrizTemporal(1, 12, 1) = FormatNumber(RegistroSQL("Soporte50"), 2, True, False, True)
   MatrizTemporal(1, 12, 2) = RegistroSQL("FechaSoporte50")
   MatrizTemporal(1, 12, 3) = FormatNumber(RegistroSQL("Resistencia50"), 2, True, False, True)
   MatrizTemporal(1, 12, 4) = RegistroSQL("FechaResistencia50")

   MatrizTemporal(1, 13, 1) = FormatNumber(RegistroSQL("Soporte200"), 2, True, False, True)
   MatrizTemporal(1, 13, 2) = RegistroSQL("FechaSoporte200")
   MatrizTemporal(1, 13, 3) = FormatNumber(RegistroSQL("Resistencia200"), 2, True, False, True)
   MatrizTemporal(1, 13, 4) = RegistroSQL("FechaResistencia200")
   
   MatrizTemporal(1, 14, 1) = RegistroSQL("Fecha1TALargo")
   MatrizTemporal(1, 14, 2) = FormatNumber(RegistroSQL("Valor1TALargo"), 2, True, False, True)
   MatrizTemporal(1, 14, 3) = RegistroSQL("Fecha2TALargo")
   MatrizTemporal(1, 14, 4) = FormatNumber(RegistroSQL("Valor2TALargo"), 2, True, False, True)
   MatrizTemporal(1, 14, 5) = FormatNumber(RegistroSQL("PorTALargo"), 4, True, False, True) & " %"
   MatrizTemporal(1, 14, 6) = FormatNumber(RegistroSQL("PorAcumuladoTALargo"), 2, True, False, True) & " %"
   MatrizTemporal(1, 14, 7) = FormatNumber(RegistroSQL("DiasTALargo"), 0, True, False, True)
   MatrizTemporal(1, 14, 8) = RegistroSQL("TipoTALargo")
   
   MatrizTemporal(1, 15, 1) = RegistroSQL("Fecha1TBLargo")
   MatrizTemporal(1, 15, 2) = FormatNumber(RegistroSQL("Valor1TBLargo"), 2, True, False, True)
   MatrizTemporal(1, 15, 3) = RegistroSQL("Fecha2TBLargo")
   MatrizTemporal(1, 15, 4) = FormatNumber(RegistroSQL("Valor2TBLargo"), 2, True, False, True)
   MatrizTemporal(1, 15, 5) = FormatNumber(RegistroSQL("PorTBLargo"), 4, True, False, True) & " %"
   MatrizTemporal(1, 15, 6) = FormatNumber(RegistroSQL("PorAcumuladoTBLargo"), 2, True, False, True) & " %"
   MatrizTemporal(1, 15, 7) = FormatNumber(RegistroSQL("DiasTBLargo"), 0, True, False, True)
   MatrizTemporal(1, 15, 8) = RegistroSQL("TipoTBLargo")

   MatrizTemporal(1, 16, 1) = RegistroSQL("Fecha1TAMedio")
   MatrizTemporal(1, 16, 2) = FormatNumber(RegistroSQL("Valor1TAMedio"), 2, True, False, True)
   MatrizTemporal(1, 16, 3) = RegistroSQL("Fecha2TAMedio")
   MatrizTemporal(1, 16, 4) = FormatNumber(RegistroSQL("Valor2TAMedio"), 2, True, False, True)
   MatrizTemporal(1, 16, 5) = FormatNumber(RegistroSQL("PorTAMedio"), 4, True, False, True) & " %"
   MatrizTemporal(1, 16, 6) = FormatNumber(RegistroSQL("PorAcumuladoTAMedio"), 2, True, False, True) & " %"
   MatrizTemporal(1, 16, 7) = FormatNumber(RegistroSQL("DiasTAMedio"), 0, True, False, True)
   MatrizTemporal(1, 16, 8) = RegistroSQL("TipoTAMedio")
   
   MatrizTemporal(1, 17, 1) = RegistroSQL("Fecha1TBMedio")
   MatrizTemporal(1, 17, 2) = FormatNumber(RegistroSQL("Valor1TBMedio"), 2, True, False, True)
   MatrizTemporal(1, 17, 3) = RegistroSQL("Fecha2TBMedio")
   MatrizTemporal(1, 17, 4) = FormatNumber(RegistroSQL("Valor2TBMedio"), 2, True, False, True)
   MatrizTemporal(1, 17, 5) = FormatNumber(RegistroSQL("PorTBMedio"), 4, True, False, True) & " %"
   MatrizTemporal(1, 17, 6) = FormatNumber(RegistroSQL("PorAcumuladoTBMedio"), 2, True, False, True) & " %"
   MatrizTemporal(1, 17, 7) = FormatNumber(RegistroSQL("DiasTBMedio"), 0, True, False, True)
   MatrizTemporal(1, 17, 8) = RegistroSQL("TipoTBMedio")

   MatrizTemporal(1, 18, 1) = RegistroSQL("Fecha1TACorto")
   MatrizTemporal(1, 18, 2) = FormatNumber(RegistroSQL("Valor1TACorto"), 2, True, False, True)
   MatrizTemporal(1, 18, 3) = RegistroSQL("Fecha2TACorto")
   MatrizTemporal(1, 18, 4) = FormatNumber(RegistroSQL("Valor2TACorto"), 2, True, False, True)
   MatrizTemporal(1, 18, 5) = FormatNumber(RegistroSQL("PorTACorto"), 4, True, False, True) & " %"
   MatrizTemporal(1, 18, 6) = FormatNumber(RegistroSQL("PorAcumuladoTACorto"), 2, True, False, True) & " %"
   MatrizTemporal(1, 18, 7) = FormatNumber(RegistroSQL("DiasTACorto"), 0, True, False, True)
   MatrizTemporal(1, 18, 8) = RegistroSQL("TipoTACorto")
   
   MatrizTemporal(1, 19, 1) = RegistroSQL("Fecha1TBCorto")
   MatrizTemporal(1, 19, 2) = FormatNumber(RegistroSQL("Valor1TBCorto"), 2, True, False, True)
   MatrizTemporal(1, 19, 3) = RegistroSQL("Fecha2TBCorto")
   MatrizTemporal(1, 19, 4) = FormatNumber(RegistroSQL("Valor2TBCorto"), 2, True, False, True)
   MatrizTemporal(1, 19, 5) = FormatNumber(RegistroSQL("PorTBCorto"), 4, True, False, True) & " %"
   MatrizTemporal(1, 19, 6) = FormatNumber(RegistroSQL("PorAcumuladoTBCorto"), 2, True, False, True) & " %"
   MatrizTemporal(1, 19, 7) = FormatNumber(RegistroSQL("DiasTBCorto"), 0, True, False, True)
   MatrizTemporal(1, 19, 8) = RegistroSQL("TipoTBCorto")

   MatrizTemporal(1, 20, 1) = FormatNumber(RegistroSQL("K"), 2, True, False, True)
   MatrizTemporal(1, 21, 1) = FormatNumber(RegistroSQL("D"), 2, True, False, True)
   MatrizTemporal(1, 22, 1) = FormatNumber(RegistroSQL("DS"), 2, True, False, True)
   MatrizTemporal(1, 23, 1) = FormatNumber(RegistroSQL("DSS"), 2, True, False, True)
   MatrizTemporal(1, 24, 1) = FormatNumber(RegistroSQL("RSI"), 2, True, False, True)
   
   
   For i = 1 To 4
   
       MatrizTemporal(1, 20, i + 1) = FormatNumber(RegistroSQL("Kn" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 21, i + 1) = FormatNumber(RegistroSQL("Dn" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 22, i + 1) = FormatNumber(RegistroSQL("DSn" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 23, i + 1) = FormatNumber(RegistroSQL("DSSn" & CStr(i)), 2, True, False, True)
       MatrizTemporal(1, 24, i + 1) = FormatNumber(RegistroSQL("RSIn" & CStr(i)), 2, True, False, True)
       
   Next

   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close

' Cerramos la conexión
ConexionSQL.Close

frmFichaValor.Caption = Titulo & " | actualizado a " & MatrizTemporal(1, 1, 2)

Load frmFichaValor

End Sub

Public Sub AnalisisTecnicoMercados_RSI(Mercado As Integer)

' Si hay valor en el cierre del valor de hace 30 sesiones
If MatrizAcciones(30, 3) <> "" Then

   ReDim MatrizRSI(30, 2) '0-para control de valor RSI, 1 alzas, 2 bajas

   For Y = 1 To 30
 
       MatrizRSI(Y, 0) = 0
       MatrizRSI(Y, 1) = 0
       MatrizRSI(Y, 2) = 0
    
   Next

   For Y = 1 To 30

       ' 1 - 14
       ' ...
       ' 10 - 23
       For Z = Y To Y + (RSI_Periodo - 1)
    
        
           ' Si el cierre de hoy menos el cierre de ayer es mayor que 0
           If (MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)) > 0 Then
        
              ' Le agregamos la subida a las alzas
              MatrizRSI(Y, 1) = CDbl(MatrizRSI(Y, 1)) + CDbl(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3))
        
           Else
        
              ' Le agregamos la bajada positivas a las bajas
              MatrizRSI(Y, 2) = CDbl(MatrizRSI(Y, 2)) + CDbl(Abs(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)))
        
           End If
        
       Next

        
       ' Calculamos la media de las alzas y las bajas
       MatrizRSI(Y, 1) = CDbl(MatrizRSI(Y, 1)) / CDbl(RSI_Periodo)
       MatrizRSI(Y, 2) = CDbl(MatrizRSI(Y, 2)) / CDbl(RSI_Periodo)
    
       ' En MatrizTemporal, solo grabamos los ultimos 5 movimientos de RSI
       If Y <= 5 Then
    
          ' Para evitar problemas de division por 0
          If MatrizRSI(Y, 1) <> 0 Then
    
             MatrizTemporal(Mercado, 24, Y) = Round((MatrizRSI(Y, 1) / (MatrizRSI(Y, 2) + MatrizRSI(Y, 1))) * 100, 4)
        
          End If
       
       End If
    
       ' Para evitar problemas de division por 0
       If MatrizRSI(Y, 1) <> 0 Then
    
          ' En MatrizRSI campo 0 guardamos el valor de RSI que nos servirá para revisar FailureSwing y Divergencia
          MatrizRSI(Y, 0) = Round((MatrizRSI(Y, 1) / (MatrizRSI(Y, 2) + MatrizRSI(Y, 1))) * 100, 4)
    
       End If
    
   Next

End If

End Sub

Public Sub AnalisisTecnicoMercados_RSI_Suavizado(Mercado As Integer)

' Si hay valor en el cierre del valor de hace 30 sesiones
If MatrizAcciones(30, 3) <> "" Then

   ReDim MatrizRSI(30, 2) '0-para control de valor RSI, 1 alzas, 2 bajas

   For Y = 1 To 30
 
       MatrizRSI(Y, 0) = 0
       MatrizRSI(Y, 1) = 0
       MatrizRSI(Y, 2) = 0
    
   Next

   For Y = 30 To 1 Step -1

       ' 1 - 14
       ' ...
       ' 10 - 23
       For Z = Y To Y + (RSI_Periodo - 1)
    
           If Y = 30 Then
           
              ' Si el cierre de hoy menos el cierre de ayer es mayor que 0
              If (MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)) > 0 Then
        
                 ' Le agregamos la subida a las alzas
                 MatrizRSI(Y, 1) = CDbl(MatrizRSI(Y, 1)) + CDbl(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3))
        
              Else
        
                 ' Le agregamos la bajada positivas a las bajas
                 MatrizRSI(Y, 2) = CDbl(MatrizRSI(Y, 2)) + CDbl(Abs(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)))
        
              End If
              
           ElseIf Z = Y Then
           
              ' Si el cierre de hoy menos el cierre de ayer es mayor que 0
              If (MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)) > 0 Then
        
                 ' Le agregamos la subida a las alzas
                 MatrizRSI(Y, 1) = CDbl(MatrizRSI(Y, 1)) + CDbl(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3))
        
              Else
        
                 ' Le agregamos la bajada positivas a las bajas
                 MatrizRSI(Y, 2) = CDbl(MatrizRSI(Y, 2)) + CDbl(Abs(MatrizAcciones(Z, 3) - MatrizAcciones(Z + 1, 3)))
        
              End If
              
            
           End If
        
       Next
         
       If Y = 30 Then
       
          ' Calculamos la media de las alzas y las bajas
          MatrizRSI(Y, 1) = CDbl(MatrizRSI(Y, 1)) / CDbl(RSI_Periodo)
          MatrizRSI(Y, 2) = CDbl(MatrizRSI(Y, 2)) / CDbl(RSI_Periodo)
          
       Else
          
          ' Calculamos la media de las alzas y las bajas
          MatrizRSI(Y, 1) = (CDbl(MatrizRSI(Y + 1, 1)) * (CDbl(RSI_Periodo) - 1) + CDbl(MatrizRSI(Y, 1))) / CDbl(RSI_Periodo)
          MatrizRSI(Y, 2) = (CDbl(MatrizRSI(Y + 1, 2)) * (CDbl(RSI_Periodo) - 1) + CDbl(MatrizRSI(Y, 2))) / CDbl(RSI_Periodo)
          
       
       End If
    
       ' En MatrizTemporal, solo grabamos los ultimos 5 movimientos de RSI
       If Y <= 5 Then
    
          ' Para evitar problemas de division por 0
          If MatrizRSI(Y, 1) <> 0 Then
    
             MatrizTemporal(Mercado, 24, Y) = Round((MatrizRSI(Y, 1) / (MatrizRSI(Y, 2) + MatrizRSI(Y, 1))) * 100, 4)
        
          End If
       
       End If
    
       ' Para evitar problemas de division por 0
       If MatrizRSI(Y, 1) <> 0 Then
    
          ' En MatrizRSI campo 0 guardamos el valor de RSI que nos servira para revisar FailureSwing y Divergencia
          MatrizRSI(Y, 0) = Round((MatrizRSI(Y, 1) / (MatrizRSI(Y, 2) + MatrizRSI(Y, 1))) * 100, 4)
    
       End If
    
   Next

End If

End Sub

Public Sub AnalisisEstrategias(Mercado As Integer)

' ESTRAGEGIAS
AnalisisEstrategia_CorreccionOrdenada (Mercado)
AnalisisEstrategia_GapApertura (Mercado)
AnalisisEstrategia_PausaHarami (Mercado)
' Marcado para que coja vela inicial media o grande
AnalisisEstrategia_PausaCasiHarami (Mercado)
' Marcado para que coja vela inicial media o grande
AnalisisEstrategia_FuegoPaja (Mercado)
AnalisisEstrategia_Exceso (Mercado)
AnalisisEstrategia_Lane (Mercado)
AnalisisEstrategia_RSI (Mercado)

' CAMBIOS DE TENDENCIA
AnalisisEstrategia_Doji (Mercado)
AnalisisEstrategia_Martillo (Mercado)
AnalisisEstrategia_EstrellaFugaz (Mercado)
AnalisisEstrategia_Harami (Mercado)
AnalisisEstrategia_Cobertura (Mercado)
AnalisisEstrategia_Penetrante (Mercado)
AnalisisEstrategia_Envolventes (Mercado)

' CAMBIOS DE TENDENCIA (Líneas de velas secundarias)
AnalisisEstrategia_Peonza (Mercado)
AnalisisEstrategia_GapTasuki (Mercado)
AnalisisEstrategia_GemelosBlancos (Mercado)
' Dicen que la apertura debe estar en en mismo nivel y nosotros lo hemos puesto exactamente en el mismo precio
AnalisisEstrategia_LineasSeparacion (Mercado)
' Dicen que la cierre debe estar en mismo nivel y nosotros lo hemos puesto exactamente en el mismo precio
AnalisisEstrategia_LineasUnion (Mercado)
AnalisisEstrategia_Puntapie (Mercado)
AnalisisEstrategia_TresRios (Mercado)

End Sub

Public Sub GuardarAvisos(Mercado As Integer, Orden As String, Estrategia As String, Comentario As String, Confirmacion As String, Procedimiento As String)

' Agregamos 1 a la cantidad de avisos
MatrizAvisos(Mercado, 0, 0) = MatrizAvisos(Mercado, 0, 0) + 1

' Introducimos en avisos la orden, estrategia, comentarios y Confirmación
MatrizAvisos(Mercado, MatrizAvisos(Mercado, 0, 0), 1) = Orden
MatrizAvisos(Mercado, MatrizAvisos(Mercado, 0, 0), 2) = Estrategia
MatrizAvisos(Mercado, MatrizAvisos(Mercado, 0, 0), 3) = Comentario
MatrizAvisos(Mercado, MatrizAvisos(Mercado, 0, 0), 4) = Confirmacion
MatrizAvisos(Mercado, MatrizAvisos(Mercado, 0, 0), 5) = Procedimiento
        

End Sub

Public Sub AnalisisEstrategia_RSI(Mercado As Integer)

' SALIDA DE ZONA RSI

' Si RSI sale de zona de sobreventa COMPRA
' Si RSI sale de zona de sobrecompra VENTA
' El stop (NO CONTROLADO), se pone cuando RSI vuelve a la zona de la que salio

If RSI_AvisoSalidaZona = True Then

' Si el valor de RSI > que RSI-1 probamos compra saliendo de zona de sobreventa
   If MatrizTemporal(Mercado, 24, 1) > MatrizTemporal(Mercado, 24, 2) Then
   
      ' Si RSI >= que la zona de Sobreventa y RSIn-1 era menor que la zona de SobreVenta
      If MatrizTemporal(Mercado, 24, 1) >= RSI_SobreVenta And MatrizTemporal(Mercado, 24, 2) < RSI_SobreVenta Then
      
         GuardarAvisos Mercado, "COMPRA", "RSI Salida de Zona", "Salida de RSI(" & CStr(Round(MatrizTemporal(Mercado, 24, 1), 2)) & ") de zona de sobreventa (" & RSI_SobreVenta & ")", "", ""
                  
      End If
   
   ' Si el valor de RSI < que RSIn-1 probamos compra saliendo de zona de sobrecompra
   ElseIf MatrizTemporal(Mercado, 24, 1) < MatrizTemporal(Mercado, 24, 2) Then
   
      ' Si RSI <= que la zona de Sobrecompra y RSIn-1 era mayor que la zona de SobreCompra
      If MatrizTemporal(Mercado, 24, 1) <= RSI_SobreCompra And MatrizTemporal(Mercado, 24, 2) > RSI_SobreCompra Then
      
         GuardarAvisos Mercado, "VENTA", "RSI Salida de Zona", "Salida de RSI (" & CStr(Round(MatrizTemporal(Mercado, 24, 1), 2)) & ") de zona de sobrecompra (" & RSI_SobreCompra & ")", "", ""
            
      End If
   
   End If

End If

' Para el control de FailureSwing
Dim MinRSI As Double
Dim MaxRSI As Double
Dim PosMinRSI As Integer
Dim PosMaxRSI As Integer

' La divergencia es FailureSwing pero con comprobación que la tendencia entre RSI
' y la cotización del valor divergen (uno es alcista y el otro bajista)
If RSI_AvisoFailureSwing = True Or RSI_AvisoDivergencia = True Then

   ' Si el valor de RSI > que RSI-1 y ambos estan por encima de zona sobreventa
   If (MatrizTemporal(Mercado, 24, 1) > MatrizTemporal(Mercado, 24, 2)) And (MatrizTemporal(Mercado, 24, 2) > RSI_SobreVenta) Then

      PosMinRSI = 0
      PosMaxRSI = 0
      
      MaxRSI = 0
       
      MinRSI = MatrizRSI(2, 0)
      
      ' Recorremos los valores de RSI desde antesdeayer
      For Y = 3 To 30
      
          ' Si el valor de RSI es menor que el Mínimo
          If MatrizRSI(Y, 0) < MinRSI Then
          
             ' Lo ponemos como nuevo mínimo y asignamos su posición
             MinRSI = MatrizRSI(Y, 0)
             PosMinRSI = Y
             
             ' Forzamos la salida de For
             Y = 30
                       
          ' Si el valor de RSI es mayor que el Maximo
          ElseIf MatrizRSI(Y, 0) > MaxRSI Then
          
             ' Lo ponemos como nuevo maximo y asignamos su posicion
             MaxRSI = MatrizRSI(Y, 0)
             PosMaxRSI = Y
             
          End If
      
      Next
      
      ' Si se han marcado un minimo y un maximo
      If PosMinRSI <> 0 And PosMaxRSI <> 0 Then
      
         ' Si el minimo se dio antes que el maximo y el minimo esta por debajo de la zona de sobreventa
         If (PosMinRSI > PosMaxRSI) And MinRSI < RSI_SobreVenta Then
            
            If RSI_AvisoFailureSwing = True Then
            
               GuardarAvisos Mercado, "COMPRA", "RSI Failure Swing", "Minimo marcado hace " & CStr(PosMinRSI - 1) & " días en " & CStr(Round(MinRSI, 4)) & " por debajo de zona sobreventa (" & RSI_SobreVenta & ")" & Chr(10) & Chr(13) & "Maximo marcado hace " & CStr(PosMaxRSI - 1) & " días en " & CStr(Round(MaxRSI, 4)) & Chr(10) & Chr(13) & "Failure Swing ayer en " & CStr(Round(MatrizRSI(2, 0), 4)) & " confirmado hoy al alza", "Revisar que RSI no baja del valor de Failure Swing o la cotización no baja del mínimo de la vela del FS (" & CStr(MatrizAcciones(2, 5)) & ")", ""
                                 
            End If
            
            If RSI_AvisoDivergencia = True Then
               
               ' Si el minimo de la cotización del día de la entrada en zona de sobreventa
               ' es mayor que el minimo de la cotización del día del Failure Swing
               If MatrizAcciones(PosMinRSI, 5) > MatrizAcciones(2, 5) Then
               
                  GuardarAvisos Mercado, "COMPRA", "RSI Divergencia", "Tendencia RSI alcista y tendencia cotización divergente porque Minimo marcado hace " & CStr(PosMinRSI - 1) & " días en " & CStr(Round(MatrizAcciones(PosMinRSI, 5), 2)) & " mayor que " & CStr(Round(MatrizAcciones(2, 5), 2)) & " de ayer", "", ""
                                            
               End If

            End If
            
         End If
      
      End If
      
   ' Si el valor de RSI < que RSI-1 y ambos estan por debajo de zona sobrecompra
   ElseIf (MatrizTemporal(Mercado, 24, 1) < MatrizTemporal(Mercado, 24, 2)) And (MatrizTemporal(Mercado, 24, 2) < RSI_SobreCompra) Then

      PosMinRSI = 0
      PosMaxRSI = 0
      
      MinRSI = 100
       
      MaxRSI = MatrizRSI(2, 0)
     
      For Y = 3 To 30
      
          If MatrizRSI(Y, 0) < MinRSI Then
          
             MinRSI = MatrizRSI(Y, 0)
             PosMinRSI = Y
                       
          ElseIf MatrizRSI(Y, 0) > MaxRSI Then
          
             MaxRSI = MatrizRSI(Y, 0)
             PosMaxRSI = Y
             
             Y = 30
             
          End If
      
      Next
      
      ' Si se han marcado un minimo y un maximo
      If PosMinRSI <> 0 And PosMaxRSI <> 0 Then
      
         ' Si el maximo se dio antes que el minimo y el maximo esta por encima de la zona de sobrecompra
         If (PosMaxRSI > PosMinRSI) And MaxRSI > RSI_SobreCompra Then
            
            If RSI_AvisoFailureSwing = True Then
            
               GuardarAvisos Mercado, "VENTA", "RSI Failure Swing", "Maximo marcado hace " & CStr(PosMaxRSI - 1) & " días en " & CStr(Round(MaxRSI, 4)) & " por encima de zona sobrecompra (" & RSI_SobreCompra & ")" & Chr(10) & Chr(13) & "Minimo marcado hace " & CStr(PosMinRSI - 1) & " días en " & CStr(Round(MinRSI, 4)) & Chr(10) & Chr(13) & "Failure Swing ayer en " & CStr(Round(MatrizRSI(2, 0), 4)) & " confirmado hoy a la baja", "Revisar que RSI no sube del valor de Failure Swing o la cotización no sube del maximo de la vela del FS (" & CStr(MatrizAcciones(2, 4)) & ")", ""
                       
            End If
            
            If RSI_AvisoDivergencia = True Then
               
               ' Si el maximo de la cotización del día de la entrada en zona de sobrecompra
               ' es menor que el maximo de la cotización del día del Failure Swing
               If MatrizAcciones(PosMaxRSI, 4) < MatrizAcciones(2, 4) Then
               
                  GuardarAvisos Mercado, "VENTA", "RSI Divergencia", "Tendencia RSI bajista y tendencia cotización divergente porque Maximo marcado hace " & CStr(PosMaxRSI - 1) & " días en " & CStr(Round(MatrizAcciones(PosMaxRSI, 4), 2)) & " menor que " & CStr(Round(MatrizAcciones(2, 4), 2)) & " de ayer", "", ""
                                           
               End If

            End If
                        
         End If
      
      End If

   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_GapApertura(Mercado As Integer)

' Tiene que haber un hueco entre el maximo de ayer y la apertura de hoy
' La diferencia entre el cierre de ayer y la apertura de hoy debe ser de un 2,5 % o más
' Un martillo antes del hueco
' Ruptura de zona de resistencias a continuación de hueco
' Indicios de cambio de tendencia al alza en vela precendente al gap.

Dim TipoGap As String
Dim NumGap As Integer
Dim VelaMartillo As Boolean
Dim MenConfirmacion As String
Dim MenProcedimiento As String

TipoGap = ""
NumGap = 0
VelaMartillo = False

If GapApertura_Aviso = True Then

   ' Para controlar Gap hoy o ayer
   For Y = 1 To 2

      ' Si la apertura Y el mínimo de hoy es mayor que el maximo de ayer
      If MatrizAcciones(Y, 2) > MatrizAcciones(Y + 1, 4) And MatrizAcciones(Y, 5) > MatrizAcciones(Y + 1, 4) Then
   
         ' Si la apertura y el mínimo de hoy es mayor al cierre de ayer, con una diferencia superior al 2,5%
         ' y el cierre de hoy es mayor que la apertura
         If MatrizAcciones(Y, 2) > (MatrizAcciones(Y + 1, 3) * 1.025) And MatrizAcciones(Y, 3) >= MatrizAcciones(Y, 2) And MatrizAcciones(Y, 5) > (MatrizAcciones(Y + 1, 3) * 1.025) Then
      
            TipoGap = "A"
            NumGap = Y - 1
   
         End If
   
      ' Si la apertura y el máximo de hoy es menor que el minimo de ayer
      ElseIf MatrizAcciones(Y, 2) < MatrizAcciones(Y + 1, 5) And MatrizAcciones(Y, 4) < MatrizAcciones(Y + 1, 5) Then
   
         ' Si la apertura y el máximo de hoy es menor al cierre de ayer, con una diferencia superior al 2,5%
         ' y el cierre de hoy es menor que la apertura
         If MatrizAcciones(Y, 2) < (MatrizAcciones(Y + 1, 3) * 0.975) And MatrizAcciones(Y, 3) <= MatrizAcciones(Y, 2) And MatrizAcciones(Y, 4) < (MatrizAcciones(Y + 1, 3) * 0.975) Then
   
            TipoGap = "B"
            NumGap = Y - 1
   
         End If
      
      End If
   
   Next
   
   ' Si existe Gap
   If TipoGap <> "" Then
   
      ' Si la vela anterior al gap es un martillo
      If Abs(MatrizAcciones(2 + NumGap, 7)) = 11 Then
      
         VelaMartillo = True
      
      End If
      
      ' Si el Gap es alcista
      If TipoGap = "A" Then
         
         MenConfirmacion = ""
            
         If NumGap = 1 Then
            
            ' Si el cierre de hoy supero el maximo de ayer
            If MatrizAcciones(1, 3) > MatrizAcciones(2, 4) Then
            
               MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de max última vela " & CStr(MatrizAcciones(2, 4)) & " "
               
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 5) & " o en " & MatrizAcciones(1, 5) & " Objetivo no existe"
                  
            Else
               
               MenConfirmacion = MenConfirmacion & "Confirmar rotura de max última vela " & CStr(MatrizAcciones(2, 4)) & " "
                  
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 5) & " Objetivo no existe"
                  
            End If
               
         Else
            
            MenConfirmacion = MenConfirmacion & "Confirmar rotura de max última vela " & CStr(MatrizAcciones(1, 4)) & " "
              
            MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 5) & " Objetivo no existe"
               
         End If
           
         If VelaMartillo = True Then MenConfirmacion = MenConfirmacion & "Martillo en vela anterior a Gap "
            
         MenConfirmacion = MenConfirmacion & "Confirmar ruptura de una zona de resistencias e indicios de cambio de tendencia al alza en vela precedente al gab"

            
         GuardarAvisos Mercado, "COMPRA", "Gap Aperura", "Gap Alcista entre maximo " & CStr(MatrizAcciones(2 + NumGap, 4)) & " de n-1 y apertura " & CStr(MatrizAcciones(1 + NumGap, 2)) & " de n, con diferencia superior al 2,5% entre cierre y apertura", MenConfirmacion, MenProcedimiento
            
      ' Si el Gab es bajista
      Else
         
         MenConfirmacion = ""
            
         If NumGap = 1 Then
            
            ' Si el cierre de hoy supero el minimo de ayer
            If MatrizAcciones(1, 3) < MatrizAcciones(2, 5) Then
            
               MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de min última vela " & CStr(MatrizAcciones(2, 5)) & " "
               
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 4) & " o en " & MatrizAcciones(1, 4) & " Objetivo no existe"
                  
            Else
               
               MenConfirmacion = MenConfirmacion & "Confirmar rotura de min última vela " & CStr(MatrizAcciones(2, 5)) & " "
                  
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 4) & " Objetivo no existe"
                  
            End If
               
         Else
            
            MenConfirmacion = MenConfirmacion & "Confirmar rotura de min última vela " & CStr(MatrizAcciones(1, 5)) & " "
              
            MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 4) & " Objetivo no existe"
               
         End If
           
         If VelaMartillo = True Then MenConfirmacion = MenConfirmacion & "Martillo en vela anterior a Gap "
            
         MenConfirmacion = MenConfirmacion & "Confirmar ruptura de una zona de soporte e indicios de cambio de tendencia a la baja en vela precedente al gab"

         'MsgBox "VENTA Gap apertura"
            
         GuardarAvisos Mercado, "VENTA", "Gap Aperura", "Gap Bajista entre minimo " & CStr(MatrizAcciones(2 + NumGap, 5)) & " de n-1 y apertura " & CStr(MatrizAcciones(1 + NumGap, 2)) & " de n, con diferencia superior al 2,5% entre cierre y apertura", MenConfirmacion, MenProcedimiento
                        
      End If
    
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_Exceso(Mercado As Integer)

' Una bajada del 10 por ciento en tres sesiones seguidas. Sugerencia propia, que el movimiento de las dos últimas sesiones sea de tres o más medias de vela de los últimos 20 días.
' Tendencia bajista.
' La presencia de una zona de soporte al nivel del mínimo de la última sesión de Trading (bajo la forma de una estructura gráfica de velas o una línea simple).
' Un volumen elevado durante la bajada.

Dim TipoExceso As String
Dim PorDiferencia As Double
Dim MediaVolumen As Double
Dim CuerpoVelas As Double
Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String


TipoExceso = ""

If Exceso_Aviso = True Then
   
   If IsDate(MatrizAcciones(4, 1)) Then
   ' Si la cotización de hoy es distinta a la cotización de hace 4 dias
   If MatrizAcciones(1, 3) <> MatrizAcciones(4, 3) Then
   
      PorDiferencia = (MatrizAcciones(1, 3) - MatrizAcciones(4, 3)) / MatrizAcciones(4, 3) * 100
          
      If PorDiferencia > 10 Then
          
         ' Si la tendencia de medio es alcista o lateral alcista
         If InStr(1, MatrizTemporal(Mercado, 16, 8), "ALCISTA", vbTextCompare) <> 0 Then
                
            TipoExceso = "A"
            
            MenComentario = "Contraataque sobre el Exceso, por bajada de " & Round(PorDiferencia, 2) & "% en los ultimos tres días es superior al -10%"
                
            ' Calculamos la media de volumen de los tres ultimos dias
            MediaVolumen = (MatrizAcciones(1, 6) + MatrizAcciones(2, 6) + MatrizAcciones(3, 6)) / 3
            
            ' Si la media calculada es mayor a la media de 20 dias + un 25%
            If MediaVolumen > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
               MenConfirmacion = "CONFIRMADO Volumen alto respecto a Media 20 días + 25% "
               
            ' Si la media calculada es mayor a la media de 50 dias
            ElseIf MediaVolumen > MatrizTemporal(Mercado, 6, 1) Then
            
               MenConfirmacion = "Medio CONFIRMADO Volumen alto respecto a Media 50 días "
               
            Else
            
               MenConfirmacion = "OJO Volumen bajo en exceso "
            
            End If
            
            MenConfirmacion = "Confirmar cercania zona Soporte " & MenConfirmacion
            
         End If
             
      ElseIf PorDiferencia < -10 Then
      
         ' Si la tendencia de medio es bajista o lateral bajista
         If InStr(1, MatrizTemporal(Mercado, 16, 8), "BAJISTA", vbTextCompare) <> 0 Then
                
            TipoExceso = "B"
                
            MenComentario = "Contraataque sobre el Exceso, por bajada de " & Round(PorDiferencia, 2) & "% en los ultimos tres días es superior al 10%"
            
            ' Calculamos la media de volumen de los tres ultimos dias
            MediaVolumen = (MatrizAcciones(1, 6) + MatrizAcciones(2, 6) + MatrizAcciones(3, 6)) / 3
            
            ' Si la media calculada es mayor a la media de 20 dias + un 25%
            If MediaVolumen > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
               MenConfirmacion = "CONFIRMADO Volumen alto respecto a Media 20 días + 25% "
               
            ' Si la media calculada es mayor a la media de 50 dias
            ElseIf MediaVolumen > MatrizTemporal(Mercado, 6, 1) Then
            
               MenConfirmacion = "Medio CONFIRMADO Volumen alto respecto a Media 50 días "
               
            Else
            
               MenConfirmacion = "OJO Volumen bajo en exceso "
            
            End If
            
            MenConfirmacion = "Confirmar cercania zona Resistencia " & MenConfirmacion
            
         End If
          
      End If
          
   End If
   
   End If
   
   If TipoExceso = "" Then
      
      
      ' Si las dos ultimas velas son negras
      If MatrizAcciones(1, 7) < 0 And MatrizAcciones(2, 7) < 0 Then
           
         ' Calculamos el tamaño del cuerpo restandole a la apertura de ayer el cierre de hoy
         CuerpoVelas = MatrizAcciones(2, 2) - MatrizAcciones(1, 3)
         
         ' Si el cuerpo de las velas es mayor a 3 veces la vela media de 50 dias
         If CuerpoVelas > (MatrizTemporal(Mercado, 9, 1) * 3) Then
         
            ' Si la tendencia de medio es bajista o lateral bajista
            If InStr(1, MatrizTemporal(Mercado, 16, 8), "BAJISTA", vbTextCompare) <> 0 Then
                
               TipoExceso = "B"
            
               MenComentario = "Contraataque sobre el Exceso, por bajada durante los dos ultimos días con cuepos tres veces más grandes que la media 50 días"
                
               ' Calculamos la media de volumen de los tres ultimos dias
               MediaVolumen = (MatrizAcciones(1, 6) + MatrizAcciones(2, 6)) / 2
            
               ' Si la media calculada es mayor a la media de 20 dias + un 25%
               If MediaVolumen > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
                  MenConfirmacion = "CONFIRMADO Volumen alto respecto a Media 20 días + 25% "
               
               ' Si la media calculada es mayor a la media de 50 dias
               ElseIf MediaVolumen > MatrizTemporal(Mercado, 6, 1) Then
            
                  MenConfirmacion = "Medio CONFIRMADO Volumen alto respecto a Media 50 días "
               
               Else
            
                  MenConfirmacion = "OJO Volumen bajo en exceso "
            
               End If
               
               MenConfirmacion = "Confirmar cercania zona Soporte " & MenConfirmacion
            
            End If
         
         End If
         
      ' Si las dos ultimas velas son blancas
      ElseIf MatrizAcciones(1, 7) > 0 And MatrizAcciones(2, 7) > 0 Then
      
         ' Calculamos el tamaño del cuerpo restandole al cierre de hoy la apertura de ayer
         CuerpoVelas = MatrizAcciones(1, 3) - MatrizAcciones(2, 2)
         
         ' Si el cuerpo de las velas es mayor a 3 veces la vela media de 50 dias
         If CuerpoVelas > (MatrizTemporal(Mercado, 9, 1) * 3) Then
         
            ' Si la tendencia de medio es alcista o lateral alcista
            If InStr(1, MatrizTemporal(Mercado, 16, 8), "ALCISTA", vbTextCompare) <> 0 Then
                
               TipoExceso = "A"
            
               MenComentario = "Contraataque sobre el Exceso, por subida durante los dos ultimos días con cuepos tres veces más grandes que la media 50 días"
                
               ' Calculamos la media de volumen de los tres ultimos dias
               MediaVolumen = (MatrizAcciones(1, 6) + MatrizAcciones(2, 6)) / 2
            
               ' Si la media calculada es mayor a la media de 20 dias + un 25%
               If MediaVolumen > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
                  MenConfirmacion = "CONFIRMADO Volumen alto respecto a Media 20 días + 25% "
               
               ' Si la media calculada es mayor a la media de 50 dias
               ElseIf MediaVolumen > MatrizTemporal(Mercado, 6, 1) Then
            
                  MenConfirmacion = "Medio CONFIRMADO Volumen alto respecto a Media 50 días "
               
               Else
            
                  MenConfirmacion = "OJO Volumen bajo en exceso "
            
               End If
            
               MenConfirmacion = "Confirmar cercania zona Resistencia " & MenConfirmacion
               
            End If
         
         End If
         
      End If
   
   End If
   
   ' Si se ha generado algun exceso
   If TipoExceso <> "" Then
   
      ' Si el exceso es bajista
      If TipoExceso = "B" Then
         
         MenProcedimiento = "Comprar si mañana abre por encima del cierre de hoy y se rompe el máximo de la primera media hora, Stop Loss en mínimo de la primera media hora o en el mínimo de hoy " & MatrizAcciones(1, 5) & " y Objetivo revisar zonas de resistencia"
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "COMPRA", "Contraataque sobre Exceso", MenComentario, MenConfirmacion, MenProcedimiento
         
      ' Si el exceso es alcista
      Else
      
         MenProcedimiento = "Vender si mañana abre por debajo del cierre de hoy y se rompe el mínimo de la primera media hora, Stop Loss en máximo de la primera media hora o en el máximo de hoy " & MatrizAcciones(1, 4) & " y Objetivo revisar zonas de soporte"
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "VENTA", "Contraataque sobre Exceso", MenComentario, MenConfirmacion, MenProcedimiento
         
      End If
   
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_FuegoPaja(Mercado As Integer)

 ' Tendencia alcista Corto (50 dias) o Muy Corto (20 dias)
 ' Una vela de gran cuerpo negro en la cotización de ayer
 ' Las dos últimas velas de la primera fase, como mínimo, tienen cada una el mínimo superior al mínimo precedente.
' La presencia de un soporte al nivel del mínimo de la vela fuego de paja, es decir, de la vela bajista con gran cuerpo de la segunda fase.
 ' Las medías móviles de 20 y 50 días al alza.
 ' Un volumen elevado cuando se rompe el máximo de la vela fuego de paja.
 ' Procede comprar cuando la cotización de la acción rompe el máximo de la vela de fuego de paja durante la sesión siguiente

Dim TipoFuegoPaja As String
Dim NumVelas As Integer
Dim MM20Favorable As Boolean
Dim MM50Favorable As Boolean
Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

TipoFuegoPaja = ""
MenComentario = ""

If FuegoPaja_Aviso = True Then

   ' Si la vela de ayer es grande y negra
   If MatrizAcciones(2, 7) < -30 Then
   
      ' Si hoy ha cerrado por encima del maximo de la vela grande y negra
      If MatrizAcciones(1, 3) > MatrizAcciones(2, 4) Then
      
         ' Si la tendencia de corto es alcista o lateral alcista
         If InStr(1, MatrizTemporal(Mercado, 18, 8), "ALCISTA", vbTextCompare) <> 0 Then
             
            TipoFuegoPaja = "B"
            
            MenComentario = "Fuego de Paja bajista, por gran vela negra ayer, sobrepasado el maximo de ayer hoy y tendencia alcista en corto 50d"
                              
         End If
         
         ' Si la tendencia de muy corto es alcista o lateral alcista
         If InStr(1, MatrizTemporal(Mercado, 26, 8), "ALCISTA", vbTextCompare) <> 0 Then
            
            TipoFuegoPaja = "B"
            
            If MenComentario = "" Then
            
               MenComentario = "Fuego de Paja bajista, por gran vela negra ayer, sobrepasado el maximo de ayer hoy y tendencia alcista en muy corto 20d"
               
            Else
            
               MenComentario = MenComentario & " y muy corto 20d"
            
            End If
            
         End If
      
      End If
   
   ' Si la vela de ayer es grande y blanca
   ElseIf MatrizAcciones(2, 7) > 30 Then
   
      ' Si hoy ha cerrado por debajo del minimo de la vela grande y blanca
      If MatrizAcciones(1, 3) < MatrizAcciones(2, 5) Then
      
         ' Si la tendencia de corto es bajista o lateral bajista
         If InStr(1, MatrizTemporal(Mercado, 18, 8), "BAJISTA", vbTextCompare) <> 0 Then
             
            TipoFuegoPaja = "A"
            
            MenComentario = "Fuego de Paja alcista, por gran vela blanca ayer, sobrepasado el minimo de ayer hoy y tendencia bajista en corto 50d"
                              
         End If
         
         ' Si la tendencia de muy corto es bajista o lateral bajista
         If InStr(1, MatrizTemporal(Mercado, 26, 8), "BAJISTA", vbTextCompare) <> 0 Then
            
            TipoFuegoPaja = "A"
            
            If MenComentario = "" Then
            
               MenComentario = "Fuego de Paja alcista, por gran vela blanca ayer, sobrepasado el minimo de ayer hoy y tendencia bajista en muy corto 20d"
               
            Else
            
               MenComentario = MenComentario & " y muy corto 20d"
            
            End If
            
         End If
      
      End If
      
   End If
   
   If TipoFuegoPaja <> "" Then
   
      NumVelas = 0
      MM20Favorable = False
      MM50Favorable = False
      
      ' Si el volumen de hoy es mayor a la media de 20 dias + un 25%
      If MatrizAcciones(1, 6) > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
         MenConfirmacion = "CONFIRMADO Volumen alto hoy respecto a la Media de 20d + 25% "
               
      ' Si el volumen de hoy es mayor a la media de 50 dias
      ElseIf MatrizAcciones(1, 6) > MatrizTemporal(Mercado, 6, 1) Then
            
         MenConfirmacion = "Medio CONFIRMADO Volumen alto hoy respecto a Media 50 días "
               
      Else
            
         MenConfirmacion = "OJO Volumen bajo hoy "
            
      End If
      
      ' Si el tipo de fuego paja es bajista
      If TipoFuegoPaja = "B" Then
         
         ' Recorremos unas cuantas velas desde el fuego de paja
         For Y = 3 To 20
             
             ' Si la vela tratada tiene un minimo superior al minimo de la vela precedente
             If MatrizAcciones(Y, 5) > MatrizAcciones(Y + 1, 5) Then
             
                NumVelas = NumVelas + 1
                
             Else
             
                Y = 20
             
             End If
         
         Next
         
         ' Si las velas con minimos superiores a los precedentes son dos o más
         If NumVelas >= 2 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO " & NumVelas & " precedentes a fuego paja con minimos superiores a los precedentes "
            
         Else
         
            MenConfirmacion = "OJO solo " & NumVelas & " precedentes a fuego paja con minimos superiores a los precedentes " & MenConfirmacion
            
         End If
         
         ' Si la media movil 20 de ayer es mayor a la de antes de ayer
         If MatrizTemporal(Mercado, 2, 3) > MatrizTemporal(Mercado, 2, 4) Then
            
            MM20Favorable = True
            
         End If
         
         ' Si la media movil 50 de ayer es mayor a la de antes de ayer
         If MatrizTemporal(Mercado, 3, 3) > MatrizTemporal(Mercado, 3, 4) Then
         
            MM50Favorable = True
            
         End If
         
         ' Si las medias moviles no son favorables
         If MM20Favorable = False And MM50Favorable = False Then
            
            MenConfirmacion = "OJO MM20 y 50 no favorables por bajistas " & MenConfirmacion
            
         ' Si las medias moviles son favorables
         ElseIf MM20Favorable = True And MM50Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM20 y 50 favorables por alcistas "
         
         ' Si la media movil de 20 es favorables
         ElseIf MM20Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM20 favorable por alcista "
         
         ' Si las media movil de 50 es favorables
         ElseIf MM50Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM50 favorable por alcista "
         
         End If
         
         ' Si la diferencia entre el minimo de ayer y la MM20 de ayer es menor al 0,05% de la cotización
         If Abs(MatrizAcciones(2, 5) - MatrizTemporal(Mercado, 2, 3)) < (MatrizAcciones(2, 5) * 0.0005) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Soporte en MM20 "
            
         End If
         
         ' Si la diferencia entre el minimo de ayer y la MM50 de ayer es menor al 0,05% de la cotización
         If Abs(MatrizAcciones(2, 5) - MatrizTemporal(Mercado, 3, 3)) < (MatrizAcciones(2, 5) * 0.0005) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Soporte en MM50 "
            
         End If
         
         ' Si la fecha de ayer coincide con la fecha del ultimo apoyo tendencia alcista 20 dias
         If MatrizAcciones(2, 1) = MatrizTemporal(Mercado, 26, 3) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Soporte en tendencia Alcista 20 dias "
            
         End If
         
         ' Si la fecha de ayer coincide con la fecha del ultimo apoyo tendencia alcista 50 dias
         If MatrizAcciones(2, 1) = MatrizTemporal(Mercado, 18, 3) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Soporte en tendencia Alcista 50 dias "
            
         End If
         
         MenProcedimiento = "Stop Loss en mínimo de vela de fuego de paja " & MatrizAcciones(2, 5) & " o de la vela de hoy " & MatrizAcciones(1, 5) & " y Objetivo revisar zonas de resistencia"

         ' Grabamos el aviso
         GuardarAvisos Mercado, "COMPRA", "Fuego de Paja", MenComentario, MenConfirmacion, MenProcedimiento
         
      ' Si el tipo de fuego paja es alcista
      Else
         
         ' Recorremos unas cuantas velas desde el fuego de paja
         For Y = 3 To 20
             
             ' Si la vela tratada tiene un maximo inferior al maximo de la vela precedente
             If MatrizAcciones(Y, 4) < MatrizAcciones(Y + 1, 4) Then
             
                NumVelas = NumVelas + 1
                
             Else
             
                Y = 20
             
             End If
         
         Next
         
         ' Si las velas con maximos inferiores a los precedentes son dos o más
         If NumVelas >= 2 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO " & NumVelas & " precedentes a fuego paja con maximos inferiores a los precedentes "
            
         Else
         
            MenConfirmacion = "OJO solo " & NumVelas & " precedentes a fuego paja con maximos inferiores a los precedentes " & MenConfirmacion
            
         End If
         
         ' Si la media movil 20 de ayer es menor a la de antes de ayer
         If MatrizTemporal(Mercado, 2, 3) < MatrizTemporal(Mercado, 2, 4) Then
            
            MM20Favorable = True
            
         End If
         
         ' Si la media movil 50 de ayer es menor a la de antes de ayer
         If MatrizTemporal(Mercado, 3, 3) < MatrizTemporal(Mercado, 3, 4) Then
         
            MM50Favorable = True
            
         End If
         
         ' Si las medias moviles no son favorables
         If MM20Favorable = False And MM50Favorable = False Then
            
            MenConfirmacion = "OJO MM20 y 50 no favorables por alcistas " & MenConfirmacion
            
         ' Si las medias moviles son favorables
         ElseIf MM20Favorable = True And MM50Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM20 y 50 favorables por bajistas "
         
         ' Si la media movil de 20 es favorables
         ElseIf MM20Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM20 favorable por bajista "
         
         ' Si las media movil de 50 es favorables
         ElseIf MM50Favorable = True Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO MM50 favorable por bajista "
         
         End If
         
         ' Si la diferencia entre el maximo de ayer y la MM20 de ayer es menor al 0,05% de la cotización
         If Abs(MatrizAcciones(2, 4) - MatrizTemporal(Mercado, 2, 3)) < (MatrizAcciones(2, 4) * 0.0005) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Resistencia en MM20 "
            
         End If
         
         ' Si la diferencia entre el maximo de ayer y la MM50 de ayer es menor al 0,05% de la cotización
         If Abs(MatrizAcciones(2, 4) - MatrizTemporal(Mercado, 3, 3)) < (MatrizAcciones(2, 4) * 0.0005) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Resistencia en MM50 "
            
         End If
         
         ' Si la fecha de ayer coincide con la fecha del ultimo apoyo tendencia bajista 20 dias
         If MatrizAcciones(2, 1) = MatrizTemporal(Mercado, 27, 3) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Resistencia en tendencia bajista 20 dias "
            
         End If
         
         ' Si la fecha de ayer coincide con la fecha del ultimo apoyo tendencia bajista 50 dias
         If MatrizAcciones(2, 1) = MatrizTemporal(Mercado, 19, 3) Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO Resistencia en tendencia Bajista 50 dias "
            
         End If
         
         MenProcedimiento = "Stop Loss en máximo de vela de fuego de paja " & MatrizAcciones(2, 4) & " o de la vela de hoy " & MatrizAcciones(1, 4) & " y Objetivo revisar zonas de soporte"

         ' Grabamos el aviso
         GuardarAvisos Mercado, "VENTA", "Fuego de Paja", MenComentario, MenConfirmacion, MenProcedimiento
            
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_Doji(Mercado As Integer)

 ' Una vela grande o mediana (preferiblemente negra) segunda de un doji
 ' Una tercera vela de gran cuerpo blanco
 ' La segunda vela tiene que tener el cuerpo por debajo de los cuerpos de la primera y tercera vela
 ' Importante que el nivel de cierre de la tercera vela sea muy alto por encima del cuerpo de la primera
 ' Importante que el volumen de la tercera sobrepase largamente el de la primera
 ' Más robusto si la segunda vela abre un hueco con sobras incluidas respecto a la primera y segunda y sus sombras

' Variante que la segunda vela sea una peonza preferiblemente blanca

Dim TipoDoji As String
Dim MaxPrimera As Double
Dim MinPrimera As Double
Dim MaxSegunda As Double
Dim MinSegunda As Double
Dim MaxTercera As Double
Dim MinTercera As Double
Dim MenConfirmacion As String

TipoDoji = ""

If Doji_Aviso = True Then

   ' Si la vela de anteayer es mediana o grande, la de ayer peonza o doji y la de hoy grande
   If Abs(MatrizAcciones(3, 7)) > 10 And Abs(MatrizAcciones(2, 7)) < 20 And Abs(MatrizAcciones(1, 7)) > 20 Then
      
      ' Si la apertura de anteayer es mayor o igual que el cierre
      If MatrizAcciones(3, 2) >= MatrizAcciones(3, 3) Then
      
         MaxPrimera = MatrizAcciones(3, 2)
         MinPrimera = MatrizAcciones(3, 3)
         
      Else
      
         MaxPrimera = MatrizAcciones(3, 3)
         MinPrimera = MatrizAcciones(3, 2)
      
      End If
      
      ' Si la apertura de ayer es mayor o igual que el cierre
      If MatrizAcciones(2, 2) >= MatrizAcciones(2, 3) Then
      
         MaxSegunda = MatrizAcciones(2, 2)
         MinSegunda = MatrizAcciones(2, 3)
         
      Else
      
         MaxSegunda = MatrizAcciones(2, 3)
         MinSegunda = MatrizAcciones(2, 2)
      
      End If
      
      ' Si la apertura de hoy es mayor o igual que el cierre
      If MatrizAcciones(1, 2) >= MatrizAcciones(1, 3) Then
      
         MaxTercera = MatrizAcciones(1, 2)
         MinTercera = MatrizAcciones(1, 3)
         
      Else
      
         MaxTercera = MatrizAcciones(1, 3)
         MinTercera = MatrizAcciones(1, 2)
      
      End If
      
      ' Si el maximo (sombras) de ayer es menor o igual al minimo de la primera
      ' y al minimo de la tercera y la tercera es un cuerpo grande y blanco
      If MatrizAcciones(2, 4) <= MatrizAcciones(3, 5) And MatrizAcciones(2, 4) <= MatrizAcciones(1, 5) And MatrizAcciones(1, 7) > 20 Then
         
         TipoDoji = "ABebe"
         
      ' Si el maximo de ayer es menor o igual al minimo de la primera
      ' y al minimo de la tercera y la tercera es un cuerpo grande y blanco
      ElseIf MaxSegunda <= MinPrimera And MaxSegunda <= MinTercera And MatrizAcciones(1, 7) > 20 Then
      
         TipoDoji = "AEstrella"
         
      ' Si el minimo (sombras) de ayer es mayor o igual al maximo de la primera
      ' y al maximo de la tercera y la tercera es un cuerpo grande y negro
      ElseIf MatrizAcciones(2, 5) >= MatrizAcciones(3, 4) And MatrizAcciones(2, 5) >= MatrizAcciones(1, 4) And MatrizAcciones(1, 7) < -20 Then
      
         TipoDoji = "BBebe"
         
      ' Si el minimo de ayer es mayor o igual al maximo de la primera
      ' y al maximo de la tercera y la tercera es un cuerpo grande y negro
      ElseIf MinSegunda >= MaxPrimera And MinSegunda >= MaxTercera And MatrizAcciones(1, 7) < -20 Then
      
         TipoDoji = "BEstrella"
      
      End If
      
   
   End If
   
   If TipoDoji <> "" Then
     
      ' Si el volumen de hoy es mayor al volumen de anteayer
      If MatrizAcciones(1, 6) > MatrizAcciones(3, 6) Then
      
         MenConfirmacion = "CONFIRMADO Más volumen hoy que anteayer "
   
      Else
      
         MenConfirmacion = "OJO el volumen de hoy es más bajo que anteayer "
      
      End If
      CierreOk = False
   
      ' Si el tipo doji es alcista
      If Left(TipoDoji, 1) = "A" Then
      
         ' Si el cierre de hoy es mayor al maximo del cuerpo de anteayer
         If MatrizAcciones(1, 3) > MaxPrimera Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cierre de hoy por encima del cuerpo de anteayer "
            
         Else
         
            MenConfirmacion = "ESPERAR Confirmación de rotura al alza del máximo del cuerpo de anteayer " & MaxPrimero & " " & MenConfirmacion
         
         End If
      
      Else
      
         ' Si el cierre de hoy es menor al minimo del cuerpo de anteayer
         If MatrizAcciones(1, 3) < MinPrimera Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cierre de hoy por debajo del cuerpo de anteayer "
            
         Else
         
            MenConfirmacion = "ESPERAR Confirmación de rotura a la baja del mínimo del cuerpo de anteayer " & MinPrimero & " " & MenConfirmacion
         
         End If
      
      End If
      
      If TipoDoji = "ABebe" Then
         
         ' Si la vela de ayer es un doji
         If Abs(MatrizAcciones(2, 7)) < 10 Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA ALCISTA", "Bebe abandonado matinal Doji", "Porque la vela de ayer es doji, su cuerpo y sombras estan por debajo de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
         
         ' Si es una peonza
         Else
            
            ' Si el cuerpo de ayer es blanco
            If MatrizAcciones(2, 7) > 0 Then
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA ALCISTA", "Bebe abandonado matinal", "Porque la vela de ayer es peonza blanca, su cuerpo y sombras estan por debajo de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            ' Si el cuerpo de ayer es negro
            Else
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA ALCISTA", "Bebe abandonado matinal", "Porque la vela de ayer es peonza negra (deberia ser blanca), su cuerpo y sombras estan por debajo de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            End If
         
         End If
         
      ElseIf TipoDoji = "AEstrella" Then
      
         ' Si la vela de ayer es un doji
         If Abs(MatrizAcciones(2, 7)) < 10 Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA ALCISTA", "Estrella de la mañana Doji", "Porque la vela de ayer es doji, su cuerpo esta por debajo de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
         
         ' Si es una peonza
         Else
            
            ' Si el cuerpo de ayer es blanco
            If MatrizAcciones(2, 7) > 0 Then
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA ALCISTA", "Estrella de la mañana", "Porque la vela de ayer es peonza blanca, su cuerpo esta por debajo de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            ' Si el cuerpo de ayer es negro
            Else
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA ALCISTA", "Estrella de la mañana", "Porque la vela de ayer es peonza negra (deberia ser blanca), su cuerpo esta por debajo de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y blanco", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            End If
         
         End If
     
      ElseIf TipoDoji = "BBebe" Then
         
         ' Si la vela de ayer es un doji
         If Abs(MatrizAcciones(2, 7)) < 10 Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA BAJISTA", "Bebe abandonado nocturno Doji", "Porque la vela de ayer es doji, su cuerpo y sombras estan por encima de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
         
         ' Si es una peonza
         Else
            
            ' Si el cuerpo de ayer es negro
            If MatrizAcciones(2, 7) < 0 Then
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA BAJISTA", "Bebe abandonado nocturno", "Porque la vela de ayer es peonza negra, su cuerpo y sombras estan por encima de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            ' Si el cuerpo de ayer es blanco
            Else
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(1) VUELTA A TENDENCIA BAJISTA", "Bebe abandonado nocturno", "Porque la vela de ayer es peonza blanca (deberia ser negra), su cuerpo y sombras estan por encima de los cuerpos y sombras de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            End If
         
         End If
         
      ElseIf TipoDoji = "BEstrella" Then
      
         ' Si la vela de ayer es un doji
         If Abs(MatrizAcciones(2, 7)) < 10 Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA BAJISTA", "Estrella nocturna Doji", "Porque la vela de ayer es doji, su cuerpo esta por encima de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
         
         ' Si es una peonza
         Else
            
            ' Si el cuerpo de ayer es negro
            If MatrizAcciones(2, 7) < 0 Then
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA BAJISTA", "Estrella nocturna", "Porque la vela de ayer es peonza negra, su cuerpo esta por encima de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"
               
            ' Si el cuerpo de ayer es blanco
            Else
            
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(2) VUELTA A TENDENCIA BAJISTA", "Estrella nocturna", "Porque la vela de ayer es peonza blanca (deberia ser negra), su cuerpo esta por encima de los cuerpos de hoy y anteayer y el cuerpo de hoy es grande y negro", MenConfirmacion, "Revisar soportes y resitencias, para Stop y Objetivo"

            End If
         
         End If
      
      End If
      
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_Martillo(Mercado As Integer)

 ' Un martillo, preferiblemente blanco y sin sombra superior
 ' Tendencia bajista.
 ' Confirmado por vela blanca al dia siguiente con cierre superior al del martillo
 ' Importante tener buen volumen en la formación del martillo
 ' La presencia de una zona de soporte en el martillo

Dim NumDia As Integer
Dim MaxMartillo As Double

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If Martillo_Aviso = True Then

   NumDia = 0

   For Y = 1 To 2
   
       ' Si la vela de ayer o de hoy es un martillo
       If Abs(MatrizAcciones(Y, 7)) = 11 Then
       
          NumDia = Y
       
       End If
   
   Next
   
   ' Si la apertura del martillo es mayor al cierre
   If MatrizAcciones(NumDia, 2) > MatrizAcciones(NumDia, 3) Then
         
      'Cojemos la apertura
      MaxMartillo = MatrizAcciones(NumDia, 2)
      
   Else
          
      'Cojemos el cierre
      MaxMartillo = MatrizAcciones(NumDia, 3)
         
   End If
   
   ' Si existio un martillo ayer
   If NumDia = 2 Then

      ' Si el cierre de hoy es mayor al maximo del martillo
      If MatrizAcciones(1, 3) > MaxMartillo Then
         
         ' Lo dejamos como esta
         NumDia = 2
            
      Else
         
         ' No le hacemos caso
         NumDia = 0
         
      End If
   
   End If
   
   
   ' Si existio un martillo hoy o ayer
   If NumDia <> 0 Then
   
      ' Si la tendencia de medio o corto es alcista o lateral bajista
      If InStr(1, MatrizTemporal(Mercado, 16, 8), "BAJISTA", vbTextCompare) <> 0 Or InStr(1, MatrizTemporal(Mercado, 18, 8), "BAJISTA", vbTextCompare) <> 0 Then
         
         ' Si el martillo es blanco
         If MatrizAcciones(NumDia, 7) = 11 Then
         
            MenConfirmacion = "Martillo blanco "
            
         Else
         
            MenConfirmacion = "OJO el Martillo NO es blanco "
         
         End If
         
         ' Si el maximo del martillo es igual al maximo del cuerpo del martillo
         If MatrizAcciones(NumDia, 4) = MaxMartillo Then
         
            MenConfirmacion = MenConfirmacion & " sin sombra superior "
            
         Else
         
            MenConfirmacion = MenConfirmacion & " OJO con sombra superior "
         
         End If
         
         
         
         ' Si el día de la aparición del martillo fue ayer
         If NumDia = 2 Then
      
            MenConfirmacion = MenConfirmacion & "CONFIRMADO hoy con cierre por encima del cuerpo del martillo de ayer "
      
         Else
      
            MenConfirmacion = MenConfirmacion & "Esperar a mañana, una gran vela blanca con cierre por encima de " & MaxMartillo & " "
      
         End If
      
         ' Si el volumen del martillo es mayor a la media de 20 dias + un 25%
         If MatrizAcciones(NumDia, 6) > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
            MenConfirmacion = MenConfirmacion & "Volumen alto respecto a Media 20 días + 25% "
               
         ' Si el volumen del martillo es mayor a la media de 50 dias
         ElseIf MatrizAcciones(NumDia, 6) > MatrizTemporal(Mercado, 6, 1) Then
            
            MenConfirmacion = MenConfirmacion & "Volumen alto respecto a Media 50 días "
               
         Else
            
            MenConfirmacion = "OJO Volumen bajo en martillo " & MenConfirmacion
            
         End If
                  
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(3) VUELTA A TENDENCIA ALCISTA", "Martillo", "Porque la vela de ayer o la de hoy son Martillo", MenConfirmacion, "Revisar que zona de aparición de martillo sea zona de soporte"
         
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_EstrellaFugaz(Mercado As Integer)

' Una estrella fugaz, preferiblemente negro y sin sombra inferior
' Tendencia alcista.
' Confirmado por vela negra al dia siguiente con cierre inferior al del la estrella
' Importante tener buen volumen en la formación de la estrella
' La presencia de una zona de resitencia en la estrella.

Dim NumDia As Integer
Dim MinEstrella As Double

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If EstrellaFugaz_Aviso = True Then

   NumDia = 0

   For Y = 1 To 2
   
       ' Si la vela de ayer o de hoy es una estrella fugaz
       If Abs(MatrizAcciones(Y, 7)) = 12 Then
       
          NumDia = Y
       
       End If
   
   Next
   
   ' Si la apertura de la estrella es menor al cierre
   If MatrizAcciones(NumDia, 2) < MatrizAcciones(NumDia, 3) Then
         
      'Cojemos la apertura
      MinEstrella = MatrizAcciones(NumDia, 2)
      
   Else
          
      'Cojemos el cierre
      MinEstrella = MatrizAcciones(NumDia, 3)
         
   End If
   
   ' Si existio una estrella ayer
   If NumDia = 2 Then

      ' Si el cierre de hoy es mayor al maximo del martillo
      If MatrizAcciones(1, 3) < MinEstrella Then
         
         ' Lo dejamos como esta
         NumDia = 2
            
      Else
         
         ' No le hacemos caso
         NumDia = 0
         
      End If
   
   End If
   
   
   ' Si existio una estrella hoy o ayer
   If NumDia <> 0 Then
   
      ' Si la tendencia de medio o corto es alcista o lateral alcista
      If InStr(1, MatrizTemporal(Mercado, 16, 8), "ALCISTA", vbTextCompare) <> 0 Or InStr(1, MatrizTemporal(Mercado, 18, 8), "ALCISTA", vbTextCompare) <> 0 Then
         
         ' Si la estrella es negra
         If MatrizAcciones(NumDia, 7) = -12 Then
         
            MenConfirmacion = "Estrella fugaz negra "
            
         Else
         
            MenConfirmacion = "OJO la Estrella fugaz NO es negra "
         
         End If
         
         ' Si el minimo de la estrella es igual al minimo del cuerpo de la estrella
         If MatrizAcciones(NumDia, 5) = MinEstrella Then
         
            MenConfirmacion = MenConfirmacion & " sin sombra inferior "
            
         Else
         
            MenConfirmacion = MenConfirmacion & " OJO con sombra inferior "
         
         End If
         
         
         
         ' Si el día de la aparición de la estrella fue ayer
         If NumDia = 2 Then
      
            MenConfirmacion = MenConfirmacion & "CONFIRMADO hoy con cierre por debajo del cuerpo de la estrella de ayer "
      
         Else
      
            MenConfirmacion = MenConfirmacion & "Esperar a mañana, una gran vela negra con cierre por debajo de " & MinEstrella & " "
      
         End If
      
         ' Si el volumen del martillo es mayor a la media de 20 dias + un 25%
         If MatrizAcciones(NumDia, 6) > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
            MenConfirmacion = MenConfirmacion & "Volumen alto respecto a Media 20 días + 25% "
               
         ' Si el volumen del martillo es mayor a la media de 50 dias
         ElseIf MatrizAcciones(NumDia, 6) > MatrizTemporal(Mercado, 6, 1) Then
            
            MenConfirmacion = MenConfirmacion & "Volumen alto respecto a Media 50 días "
               
         Else
            
            MenConfirmacion = "OJO Volumen bajo en estrella " & MenConfirmacion
            
         End If
                  
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(3) VUELTA A TENDENCIA BAJISTA", "Estrella Fugaz", "Porque la vela de ayer o la de hoy es una Estrella Fugaz", MenConfirmacion, "Revisar que zona de aparición de la estrella sea zona de resistencia"
         
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_Harami(Mercado As Integer)

' Una gran vela, seguida de una peonza en posicion harami
' El cuerpo de la primera vela engloba el cuerpo de la segunda.
' El maximo de la primera vela esta por encima de resitencia de corto

Dim TipoHarami As String

TipoHarami = ""


If Harami_Aviso = True Then

   ' Si la vela de ayer es media o grande
   If Abs(MatrizAcciones(2, 7)) > 20 Then
   
      ' Si la vela de ayer contine en posición harami a la vela de hoy
      If MatrizAcciones(2, 3) > MatrizAcciones(1, 2) And MatrizAcciones(2, 3) > MatrizAcciones(1, 3) And MatrizAcciones(2, 2) < MatrizAcciones(1, 2) And MatrizAcciones(2, 2) < MatrizAcciones(1, 3) Then
        
         ' Si el maximo de ayer es mayor o igual que la resistencia de 20 dias
         If CDbl(MatrizAcciones(2, 4)) >= CDbl(MatrizTemporal(Mercado, 11, 3)) Then
         
            TipoHarami = "B"
         
         ' Si el minimo de ayer es menor o igual que al soporte de 20 dias
         ElseIf CDbl(MatrizAcciones(2, 5)) <= CDbl(MatrizTemporal(Mercado, 11, 1)) Then
         
            TipoHarami = "A"
            
         End If
         
       End If
       
   End If
   
   ' Si el tipo de Harami es alcista
   If TipoHarami = "A" Then
      
      ' Grabamos el aviso
      GuardarAvisos Mercado, "(4) VUELTA A TENDENCIA ALCISTA", "Harami", "Porque el cuerpo de la vela de ayer, contiene al cuerpo de la vela de hoy", "Minimo de la vela de ayer por debajo o al nivel del soporte de 20 días", ""
         
   ' Si el tipo de Harami es bajista
   ElseIf TipoHarami = "B" Then
      
      ' Grabamos el aviso
      GuardarAvisos Mercado, "(4) VUELTA A TENDENCIA BAJISTA", "Harami", "Porque el cuerpo de la vela de ayer, contiene al cuerpo de la vela de hoy", "Maximo de la vela de ayer por encima o al nivel del resistencia de 20 días", ""
   
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_Cobertura(Mercado As Integer)

 ' Primero una gran vela blanca y segundo una gran vela negra.
 ' La apertura de la vela negra esta por encima del cierre de la blanca.
 ' El cierre de la vela negra debe ser inferior a la zona central del cuerpo de la vela blanca.
 ' El máximo de la vela negra >= a la resistencia de 20 días.
 ' Señal bajista más fuerte el cierre de la vela negra, más cerca este de la apertura de la blanca.
 ' Señal bajista más fuerte cuando la mecha inferior de la vela negra sea más pequeña.
 ' Señal bajista más fuerte cuando la apertura de la vela negra, por encima de máximo vela blanca.
' Señal bajista más fuerte cuando más grandes son las velas y el volumen más elevado.
 ' Importante que el movimiento alcista previo sea precipitado y eufórico.
 ' Confirmación si la vela negra se forma con excesivo volumen.
 ' Confirmación cuando en una las dos siguientes sesiones, se cierre por debajo del mínimo de la vela negra.
 ' Stop Loss en máximo de la vela negra.

Dim NumCobertura As Integer

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

NumCobertura = 0

If Cobertura_Aviso = True Then

   For Y = 1 To 3
   
       ' Si la vela de hoy es media o grande negra y la vela de ayer es media o grande blanca
       If MatrizAcciones(Y, 7) < -20 And MatrizAcciones(Y + 1, 7) > 20 Then
          
          ' Si la apertura de la vela negra de hoy esta por encima del cierre de la blanca de ayer
          If MatrizAcciones(Y, 2) > MatrizAcciones(Y + 1, 3) Then
          
             ' Si el cierre de la vela negra de hoy esta por debajo de la zona media del cuerpo de la vela blanca de ayer
             If MatrizAcciones(Y, 3) < (((MatrizAcciones(Y + 1, 3) - MatrizAcciones(Y + 1, 2)) / 2) + MatrizAcciones(Y + 1, 2)) Then
             
                ' Si el máximo de la vela negra de hoy es mayor o igual a la resistencia de 20 días
                If CDbl(MatrizAcciones(Y, 4)) >= CDbl(MatrizTemporal(Mercado, 11, 3)) Then
                   
                   ' Si la cobertura se produjo ayer o anteayer
                   If Y > 1 Then
                   
                      ' Si el cierre de hoy es menor al mínimo de la vela negra
                      If MatrizAcciones(1, 3) < MatrizAcciones(Y, 5) Then
             
                         NumCobertura = Y
                         
                         Y = 4
                         
                      End If
                   
                   Else
                   
                      NumCobertura = Y
       
                      Y = 4
                   
                   End If
                   
                End If
                
             End If
          
          End If
          
       End If
   
   Next
   
   ' Si existe cobertura
   If NumCobertura <> 0 Then
      
      MenComentario = ""
      
      ' Si el cierre de la vela negra esta en la cuarta parte de abajo de la vela blanca
      If MatrizAcciones(NumCobertura, 3) < (((MatrizAcciones(NumCobertura + 1, 3) - MatrizAcciones(NumCobertura + 1, 2)) / 4) + MatrizAcciones(NumCobertura + 1, 2)) Then
             
         MenComentario = MenComentario & "FUERZA Cierre de la vela negra de la cobertura, en la cuarta parte inferior de la vela blanca "
             
      End If
      
      ' Si la mecha inferior tiene un tamaño menor al 10% del cuerpo de la vela negra
      If Abs(MatrizAcciones(NumCobertura, 3) - MatrizAcciones(NumCobertura, 5)) < (Abs(MatrizAcciones(NumCobertura, 2) - MatrizAcciones(NumCobertura, 3)) * 0.1) Then
             
         MenComentario = MenComentario & "FUERZA Mecha inferior de la vela negra muy pequeña "
             
      End If
      
      ' Si la apertura de la vela negra esta por encima del maximo de la vela blanca
      If MatrizAcciones(NumCobertura, 2) > MatrizAcciones(NumCobertura + 1, 4) Then
             
         MenComentario = MenComentario & "FUERZA Apertura de vela negra por encima del máximo de la blanca "
             
      End If
      
      ' Si los cuerpos de la vela negra y de la vela blanca son grandes
      If Abs(MatrizAcciones(NumCobertura, 7)) > 30 And Abs(MatrizAcciones(NumCobertura + 1, 7)) > 30 Then
             
         MenComentario = MenComentario & "FUERZA Velas grandes "
             
      End If
      
      ' Si el volumen de la cobertura es mayor a la media de 20 dias + un 25%
      If (MatrizAcciones(NumCobertura, 6) + MatrizAcciones(NumCobertura + 1, 6)) > (MatrizTemporal(Mercado, 5, 1) * 2 * 1.25) Then
               
          MenComentario = MenComentario & "Volumen alto en cobertura, respecto a Media 20 días + 25% "
  
      End If
      
      ' Si se ha producido hoy
      If NumCobertura = 1 Then
      
         MenConfirmacion = "Esperar confirmación con cierre por debajo del mínimo de hoy " & MatrizAcciones(1, 5) & " "
      
      ' Si se ha producido ayer
      ElseIf NumCobertura = 2 Then
      
         MenConfirmacion = "CONFIRMADA por cierre de hoy por debajo del mínimo de ayer "
            
      ' Si se ha producido anteayer
      ElseIf NumCobertura = 2 Then
      
         MenConfirmacion = "CONFIRMADA por cierre de hoy por debajo del mínimo de anteayer "
         
      End If

   
      ' Si el volumen de la vela negra es mayor a la media de 20 dias + un 25%
      If MatrizAcciones(NumCobertura, 6) > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
            MenConfirmacion = MenConfirmacion & "Volumen alto en vela negra cobertura, respecto a Media 20 días + 25% "
  
      End If
      
      MenProcedimiento = "Stop Loss en máximo vela negra de la cobertura " & MatrizAcciones(NumCobertura, 4) & " Importante que el movimiento alcista previo sea precipitado y eufórico"
   
   
      ' Grabamos el aviso
      GuardarAvisos Mercado, "(5) VUELTA A TENDENCIA BAJISTA", "Cobertura de nube negra", MenComentario, MenConfirmacion, MenProcedimiento
         
   End If

   
End If

End Sub

Public Sub AnalisisEstrategia_Penetrante(Mercado As Integer)

' Primero una gran vela negra y segundo una gran vela blanca.
' La apertura de la vela blanca esta por debajo del cierre de la negra.
' El cierre de la vela blanca debe ser superior a la zona central del cuerpo de la vela negra.
' El mínimo de la vela blanca <= al soporte de 20 días.
' Señal alcista más fuerte el cierre de la vela blanca, más cerca este de la apertura de la negra.
' Señal alcista más fuerte cuando la mecha superior de la vela blanca sea más pequeña.
' Señal alcista más fuerte cuando la apertura de la vela blanca, por debajo del mínimo vela negra.
' Señal alcista más fuerte cuando más grandes son las velas y el volumen más elevado.
' Importante que el movimiento bajista previo sea precipitado y eufórico.
' Confirmación si la vela blanca se forma con excesivo volumen.
' Confirmación cuando en una las dos siguientes sesiones, se cierre por encima del máximo de la vela blanca.
' Stop Loss en mínimo de la vela blanca.

Dim NumPenetrante As Integer

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

NumPenetrante = 0

If Penetrante_Aviso = True Then

   For Y = 1 To 3
   
       ' Si la vela de hoy es media o grande blanca y la vela de ayer es media o grande negra
       If MatrizAcciones(Y, 7) > 20 And MatrizAcciones(Y + 1, 7) < -20 Then
          
          ' Si la apertura de la vela blanca de hoy esta por debajo del cierre de la negra de ayer
          If MatrizAcciones(Y, 2) < MatrizAcciones(Y + 1, 3) Then
          
             ' Si el cierre de la vela blanca de hoy esta por encima de la zona media del cuerpo de la vela negra de ayer
             If MatrizAcciones(Y, 3) > (((MatrizAcciones(Y + 1, 2) - MatrizAcciones(Y + 1, 3)) / 2) + MatrizAcciones(Y + 1, 3)) Then
             
                ' Si el mínimo de la vela blanca de hoy es menor o igual al soporte de 20 días
                If CDbl(MatrizAcciones(Y, 5)) <= CDbl(MatrizTemporal(Mercado, 11, 1)) Then
                   
                   ' Si la penetrante se produjo ayer o anteayer
                   If Y > 1 Then
                   
                      ' Si el cierre de hoy es mayor al máximo de la vela blanca
                      If MatrizAcciones(1, 3) > MatrizAcciones(Y, 4) Then
             
                         NumPenetrante = Y
                         
                         Y = 4
                         
                      End If
                   
                   Else
                   
                      NumPenetrante = Y
       
                      Y = 4
                   
                   End If
                   
                End If
                
             End If
          
          End If
          
       End If
   
   Next
   
   ' Si existe cobertura
   If NumPenetrante <> 0 Then
      
      MenComentario = ""
      
      ' Si el cierre de la vela blanca esta en la cuarta parte de arriba de la vela negra
      If MatrizAcciones(NumPenetrante, 3) > (MatrizAcciones(NumPenetrante + 1, 2) - ((MatrizAcciones(NumPenetrante + 1, 2) - MatrizAcciones(NumPenetrante + 1, 3)) / 4)) Then
             
         MenComentario = MenComentario & "FUERZA Cierre de la vela blanca de la penetrante, en la cuarta parte superior de la vela negra "
             
      End If
      
      ' Si la mecha superior tiene un tamaño menor al 10% del cuerpo de la vela blanca
      If Abs(MatrizAcciones(NumPenetrante, 4) - MatrizAcciones(NumPenetrante, 3)) < (Abs(MatrizAcciones(NumPenetrante, 3) - MatrizAcciones(NumPenetrante, 2)) * 0.1) Then
             
         MenComentario = MenComentario & "FUERZA Mecha superior de la vela blanca muy pequeña "
             
      End If
      
      ' Si la apertura de la vela blanca esta por debajo del mínimo de la vela negra
      If MatrizAcciones(NumPenetrante, 2) < MatrizAcciones(NumPenetrante + 1, 5) Then
             
         MenComentario = MenComentario & "FUERZA Apertura de vela blanca por debajo del mínimo de la negra "
             
      End If
      
      ' Si los cuerpos de la vela blanca y de la vela negra son grandes
      If Abs(MatrizAcciones(NumPenetrante, 7)) > 30 And Abs(MatrizAcciones(NumPenetrante + 1, 7)) > 30 Then
             
         MenComentario = MenComentario & "FUERZA Velas grandes "
             
      End If
      
      ' Si el volumen de la penetrante es mayor a la media de 20 dias + un 25%
      If (MatrizAcciones(NumPenetrante, 6) + MatrizAcciones(NumPenetrante + 1, 6)) > (MatrizTemporal(Mercado, 5, 1) * 2 * 1.25) Then
               
          MenComentario = MenComentario & "Volumen alto en penetrante, respecto a Media 20 días + 25% "
  
      End If
      
      ' Si se ha producido hoy
      If NumPenetrante = 1 Then
      
         MenConfirmacion = "Esperar confirmación con cierre por encima del máximo de hoy " & MatrizAcciones(1, 4) & " "
      
      ' Si se ha producido ayer
      ElseIf NumPenetrante = 2 Then
      
         MenConfirmacion = "CONFIRMADA por cierre de hoy por encima del máximo de ayer "
            
      ' Si se ha producido anteayer
      ElseIf NumPenetrante = 2 Then
      
         MenConfirmacion = "CONFIRMADA por cierre de hoy por encima del máximo de anteayer "
         
      End If

   
      ' Si el volumen de la vela blanca es mayor a la media de 20 dias + un 25%
      If MatrizAcciones(NumPenetrante, 6) > (MatrizTemporal(Mercado, 5, 1) * 1.25) Then
               
            MenConfirmacion = MenConfirmacion & "Volumen alto en vela blanca penetrante, respecto a Media 20 días + 25% "
  
      End If
      
      MenProcedimiento = "Stop Loss en mínimo vela blanca de la Penetrante " & MatrizAcciones(NumPenetrante, 5) & " Importante que el movimiento bajista previo sea precipitado y eufórico"
   
   
      ' Grabamos el aviso
      GuardarAvisos Mercado, "(5) VUELTA A TENDENCIA ALCISTA", "Penetrante", MenComentario, MenConfirmacion, MenProcedimiento
         
   End If

   
End If

End Sub

Public Sub AnalisisEstrategia_Envolventes(Mercado As Integer)

' Primera vela negra y segunda blanca.
' El cuerpo blanco engloba el cuerpo de la vela negra.
' Importante si el cuerpo blanco engloba el cuerpo y las sombras de la vela negra es una Key reversal alcista.
' Importante si la vela negra es pequeña y la blanca muy grande.
' A tener en cuenta volumen en vela negra no es muy fuerte.
' Confirmación vuelta a tendenia alcista si el tercer día se produce un cierre por encima de la mitad de la vela blanca.


Dim NumEnvolvente As Integer
Dim TipoEnvolvente As String
Dim KeyReversal As Boolean
Dim CuerposVelas As Boolean
Dim VolumenVela As Boolean

Dim MenComentario As String
Dim MenConfirmacion As String

If Envolventes_Aviso = True Then
   
   NumEnvolvente = 0
   KeyReversal = False
   CuerposVelas = False
   VolumenVela = False

   For Y = 1 To 2
   
       ' Si la vela de hoy es blanca y la vela de ayer es negra
       If MatrizAcciones(Y, 7) > 0 And MatrizAcciones(Y + 1, 7) < 0 Then
       
          ' Si el cierre del la vela blanca es mayor a la apertura de ayer y la apertura de hoy es menor al cierre de ayer, es envolvente
          If MatrizAcciones(Y, 3) > MatrizAcciones(Y + 1, 2) And MatrizAcciones(Y, 2) < MatrizAcciones(Y + 1, 3) Then
           
             ' Si el cierre del la vela blanca es mayor al maximo de ayer y la apertura de hoy es menor al minimo de ayer, es key reversal
             If MatrizAcciones(Y, 3) > MatrizAcciones(Y + 1, 4) And MatrizAcciones(Y, 2) < MatrizAcciones(Y + 1, 5) Then
           
                KeyReversal = True
                
             End If
             
             ' Si la vela de hoy es grande blanca y la de ayer es negra media o pequeña
             If MatrizAcciones(Y, 7) > 30 And MatrizAcciones(Y + 1, 7) < -20 Then
             
                CuerposVelas = True
                
             End If
             
             ' Si el volumen de la vela negra es menor a la media de 20 dias
             If MatrizAcciones(Y + 1, 6) < MatrizTemporal(Mercado, 5, 1) Then
      
                VolumenVela = True
                
             End If
             
             TipoEnvolvente = "A"
             
             ' Si la envolvente se dio ayer
             If Y = 2 Then
                
                ' Si el cierre de hoy queda por encima de la mitad del cuerpo de la vela blanca
                If MatrizAcciones(1, 3) > (((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)) * 0.5) + MatrizAcciones(Y, 2)) Then
                
                   NumEnvolvente = Y
                   
                   Y = 3
                   
                End If
             
             Else
             
                NumEnvolvente = Y
                   
                Y = 3
             
             End If
             
          End If
          
       ' Si la vela de hoy es negra y la vela de ayer es blanca
       ElseIf MatrizAcciones(Y, 7) < 0 And MatrizAcciones(Y + 1, 7) > 0 Then
       
          ' Si el cierre del la vela negra es mayor a la apertura de ayer y la apertura de hoy es menor al cierre de ayer, es envolvente
          If MatrizAcciones(Y, 3) < MatrizAcciones(Y + 1, 2) And MatrizAcciones(Y, 2) > MatrizAcciones(Y + 1, 3) Then
           
             ' Si el cierre del la vela megra es menor al minimo de ayer y la apertura de hoy es mayor al maximo de ayer, es key reversal
             If MatrizAcciones(Y, 3) < MatrizAcciones(Y + 1, 5) And MatrizAcciones(Y, 2) > MatrizAcciones(Y + 1, 4) Then
           
                KeyReversal = True
                
             End If
             
             ' Si la vela de hoy es grande negra y la de ayer es blanca media o pequeña
             If MatrizAcciones(Y, 7) < -30 And MatrizAcciones(Y + 1, 7) > 20 Then
             
                CuerposVelas = True
                
             End If
             
             ' Si el volumen de la vela blanca es menor a la media de 20 dias
             If MatrizAcciones(Y + 1, 6) < MatrizTemporal(Mercado, 5, 1) Then
      
                VolumenVela = True
                
             End If
             
             TipoEnvolvente = "B"
             
             ' Si la envolvente se dio ayer
             If Y = 2 Then
                
                ' Si el cierre de hoy queda por debajo de la mitad del cuerpo de la vela negra
                If MatrizAcciones(1, 3) < (((MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3)) * 0.5) + MatrizAcciones(Y, 3)) Then
                
                   NumEnvolvente = Y
                   
                   Y = 3
                   
                End If
             
             Else
             
                NumEnvolvente = Y
                   
                Y = 3
             
             End If
             
          End If
          
       End If
       
   Next
   
   
   
   ' Si existe Envolvente
   If NumEnvolvente <> 0 Then
      
      ' Si la envolvente se ha dado hoy
      If NumEnvolvente = 1 Then
      
         If TipoEnvolvente = "A" Then
         
            MenComentario = "Engullimiento alcista, porque el cuerpo de la vela blanca de hoy engulle el cuerpo de la vela negra de ayer "
            MenConfirmacion = "Esperar confirmación mañana con cierre superior a " & (((MatrizAcciones(NumEnvolvente, 3) - MatrizAcciones(NumEnvolvente, 2)) * 0.5) + MatrizAcciones(NumEnvolvente, 2)) & " mitad del cuerpo blanco "
            
         Else
         
            MenComentario = "Engullimiento bajista, porque el cuerpo de la vela negra de hoy engulle el cuerpo de la vela blanca de ayer "
            MenConfirmacion = "Esperar confirmación mañana con cierre inferior a " & (((MatrizAcciones(NumEnvolvente, 2) - MatrizAcciones(NumEnvolvente, 3)) * 0.5) + MatrizAcciones(NumEnvolvente, 3)) & " mitad del cuerpo negro "
            
         End If
      
      ' Si la envolvente se dio ayer
      Else
      
         If TipoEnvolvente = "A" Then
         
            MenComentario = "Engullimiento alcista, porque el cuerpo de la vela blanca de ayer engulle el cuerpo de la vela negra de anteayer "
            MenConfirmacion = "CONFIRMADO hoy con cierre superior a la mitad del cuerpo blanco "
            
         Else
         
            MenComentario = "Engullimiento bajista, porque el cuerpo de la vela negra de ayer engulle el cuerpo de la vela blanca de anteayer "
            MenConfirmacion = "CONFIRMADO hoy con cierre inferior a la mitad del cuerpo negro "
            
         End If
      
      End If
      
      'If KeyReversal = True Then MenComentario = "KEY REVERSAL " & MenComentario
      
      If CuerposVelas = True Then MenComentario = MenComentario & ", Tamaños de velas propicios, primera pequeña y segunda grande "
      
      If VolumenVela = True Then MenComentario = MenComentario & ", Bueno porque el volumen de la primera vela de la envolvente es bajo "
   
      
      If TipoEnvolvente = "A" Then
      
         If KeyReversal = True Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(6) VUELTA A TENDENCIA ALCISTA", "Key Reversal Alcista", MenComentario, MenConfirmacion, ""
         
         Else
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(7) VUELTA A TENDENCIA ALCISTA", "Envolvente Alcista", MenComentario, MenConfirmacion, ""
          
         End If
          
      Else
      
         If KeyReversal = True Then
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(6) VUELTA A TENDENCIA BAJISTA", "Key Reversal Bajista", MenComentario, MenConfirmacion, ""
         
         Else
         
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(7) VUELTA A TENDENCIA BAJISTA", "Envolvente Bajista", MenComentario, MenConfirmacion, ""
          
         End If
         
      End If
      
   End If

End If

End Sub

Public Sub AnalisisEstrategia_Peonza(Mercado As Integer)

 ' Peonza negra que marca resistencia de 20 días.
 ' Sobras más grandes que el cuerpo.
 ' Máximo superior al de la vela anterior.
 ' El cuerpo de la peonza esta con la apertura por encima del cierre de la anterior y el cierre por debajo del cierre de la anterior.
' Avisar de revisión de tendencia alcista significativa antes de la peonza.
 ' Confirmación con cierre a nivel inferior de la peonza

Dim NumPeonza As Integer
Dim TipoPeonza As String
Dim MaxCuerpo As Double

Dim MenComentario As String
Dim MenConfirmacion As String

If Peonza_Aviso = True Then
   
   NumPeonza = 0
   
   ' Para tratar actuales y confirmadas
   For Y = 1 To 2
       
       ' Si la vela de hoy es peonza negra y el máximo de la peonza es mayor o igual a la resistencia de 20 días
       If MatrizAcciones(Y, 7) = -13 And CDbl(MatrizAcciones(Y, 4)) >= CDbl(MatrizTemporal(Mercado, 11, 3)) Then
           
          ' Si el maximo menos la apertura es mayor al cuerpo de la peonza y el cierre menos el minimo es mayor al cuerpo de la peonza
          If ((MatrizAcciones(Y, 4) - MatrizAcciones(Y, 2)) > (MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3))) And ((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 5)) > (MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3))) Then
       
             ' Si el maximo de la peonza es superior al maximo de la vela anterior
             If MatrizAcciones(Y, 4) > MatrizAcciones(Y + 1, 4) Then
                
                ' Si la apertura de ayer es mayor al cierre de ayer
                If MatrizAcciones(Y + 1, 2) > MatrizAcciones(Y + 1, 3) Then
                
                   MaxCuerpo = MatrizAcciones(Y + 1, 2)
                   
                Else
                
                   MaxCuerpo = MatrizAcciones(Y + 1, 3)
                   
                End If
                
                ' Si la apertura de hoy es menor al la parte superior del cuerpo de ayer y el cierre mayor al mismo
                If MatrizAcciones(Y, 2) < MaxCuerpo And MatrizAcciones(Y, 3) > MaxCuerpo Then
                   
                   ' Si la peonza se dio ayer
                   If Y = 2 Then
                   
                      ' Si el cierre de hoy es menor al cierre de la peonza negra
                      If MatrizAcciones(1, 3) < MatrizAcciones(Y, 3) Then
                      
                         NumPeonza = Y
                         TipoPeonza = "B"
                         
                         MenComentario = "Peonza negra ayer, con sombras más grandes que cuerpo, máximo superior al de la vela de anteayer, con cuerpo entre la zona superior de la vela de anteayer y en zona de resistencia"
                         MenConfirmacion = "Confirmado hoy con cierre por debajo del cierre de la peonza " & MatrizAcciones(Y, 3) & " y revisar si la tendencia anterior a la peonza es significativamente alcista"
                         
                      End If
                   
                   ' Si la peonza se ha dado hoy
                   Else
                                         
                      NumPeonza = Y
                      TipoPeonza = "B"
                      
                      MenComentario = "Peonza negra hoy, con sombras más grandes que cuerpo, máximo superior al de la vela de ayer, con cuerpo entre la zona superior de la vela de ayer y en zona de resistencia"
                      MenConfirmacion = "Esperar a cierre por debajo del cierre de la peonza " & MatrizAcciones(Y, 3) & " y revisar si la tendencia anterior a la peonza es significativamente alcista"
                      
                   End If
                   
                End If
            
             End If
          
          End If
       
       ' Si la vela de hoy es peonza blanca y el máximo de la peonza es menor o igual al soporte de 20 días
       ElseIf MatrizAcciones(Y, 7) = 13 And CDbl(MatrizAcciones(Y, 5)) <= CDbl(MatrizTemporal(Mercado, 11, 1)) Then
           
          ' Si el maximo menos el cierre es mayor al cuerpo de la peonza y la apertura menos el minimo es mayor al cuerpo de la peonza
          If ((MatrizAcciones(Y, 4) - MatrizAcciones(Y, 3)) > (MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) And ((MatrizAcciones(Y, 2) - MatrizAcciones(Y, 5)) > (MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) Then
       
             ' Si el mínimo de la peonza es menor al mínimo de la vela anterior
             If MatrizAcciones(Y, 5) < MatrizAcciones(Y + 1, 5) Then
                
                ' Si la apertura de ayer es mayor al cierre de ayer
                If MatrizAcciones(Y + 1, 2) < MatrizAcciones(Y + 1, 3) Then
                
                   MaxCuerpo = MatrizAcciones(Y + 1, 2)
                   
                Else
                
                   MaxCuerpo = MatrizAcciones(Y + 1, 3)
                   
                End If
                
                ' Si la apertura de hoy es menor al la parte inferior del cuerpo de ayer y el cierre mayor al mismo
                If MatrizAcciones(Y, 2) < MaxCuerpo And MatrizAcciones(Y, 3) > MaxCuerpo Then
                   
                   ' Si la peonza se dio ayer
                   If Y = 2 Then
                   
                      ' Si el cierre de hoy es mayor al cierre de la peonza blanca
                      If MatrizAcciones(1, 3) > MatrizAcciones(Y, 3) Then
                      
                         NumPeonza = Y
                         TipoPeonza = "A"
                         
                         MenComentario = "Peonza blanca ayer, con sombras más grandes que cuerpo, mínimo inferior al de la vela de anteayer, con cuerpo entre la zona inferior de la vela de anteayer y en zona de soporte"
                         MenConfirmacion = "Confirmado hoy con cierre por encima del cierre de la peonza " & MatrizAcciones(Y, 3) & " y revisar si la tendencia anterior a la peonza es significativamente bajista"
                         
                      End If
                   
                   ' Si la peonza se ha dado hoy
                   Else
                                         
                      NumPeonza = Y
                      TipoPeonza = "A"
                      
                      MenComentario = "Peonza blanca hoy, con sombras más grandes que cuerpo, mínimo inferior al de la vela de ayer, con cuerpo entre la zona inferior de la vela de ayer y en zona de soporte"
                      MenConfirmacion = "Esperar a cierre por encima del cierre de la peonza " & MatrizAcciones(Y, 3) & " y revisar si la tendencia anterior a la peonza es significativamente bajista"
                      
                   End If
                   
                End If
            
             End If
          
          End If
          
       End If
       
   Next
   
   
   
   ' Si existe Peonza
   If NumPeonza <> 0 Then
      
      If TipoPeonza = "A" Then
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(8) VUELTA A TENDENCIA ALCISTA", "Suelo en peonza blanca", MenComentario, MenConfirmacion, ""
          
      Else
      
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(8) VUELTA A TENDENCIA BAJISTA", "Techo en peonza negra", MenComentario, MenConfirmacion, ""
         
      End If
      
   End If

End If

End Sub

Public Sub AnalisisEstrategia_GapTasuki(Mercado As Integer)

 ' Tres velas dos primeras blancas entre las que hay un gap ascendente.
 ' Hueco alcista entre máximo primera vela y mínimo de la segunda.
 ' Tercera vela negra abre dentro del cuerpo de la segunda y cierra dentro del gap sin rellenarlo.
 ' Compra arriesgada, en el momento del cierre de la tercera vela, cuando se ve que el gap no podrá ser cerrado.
' Compra normal cuando el siguiente día se rompe el máximo de la primera hora.
' Compra conservadora, esperar a que el cierre de la sesión siguiente sea superior al de la vela negra.
' Solo hay que hacer caso al gap tasuki si es el primero de un movimiento alcista y no hacer caso cuando se hallan producido varios gaps.


Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If GapTasuki_Aviso = True Then
   
   ' Si la vela de ayer y anteayer son blancas y la de hoy es negra
   If MatrizAcciones(3, 7) > 10 And MatrizAcciones(2, 7) > 10 And MatrizAcciones(1, 7) < -10 Then
   
      ' Si el minimo de la vela de ayer es mayor al maximo de la vela de anteayer
      If MatrizAcciones(2, 5) > MatrizAcciones(3, 4) Then
      
         ' Si la apertura de hoy esta dentro del cuerpo de la vela de ayer
         If MatrizAcciones(1, 2) < MatrizAcciones(2, 3) And MatrizAcciones(1, 2) > MatrizAcciones(2, 2) Then
      
            ' Si el cierre de hoy esta en el hueco sin cerrarlo (por debajo del minimo de ayer y por encima del maximo de anteayer)
            If MatrizAcciones(1, 3) < MatrizAcciones(2, 5) And MatrizAcciones(1, 3) > MatrizAcciones(3, 4) Then
               
               MenComentario = "Hueco en velas blancas de ayer y anteayer, con vela negra hoy que abrio en el cuerpo de la vela de ayer y cierra en el hueco sin cerrarlo"
               MenConfirmacion = "Confirmación arriesgada por la formación en si, siempre que el heuco sea el primiero del movimiento alcista"
               MenProcedimiento = "Comprar si se rompe mañana el máximo de la primera hora o si el cierre de mañana es superior a la vela de hoy " & MatrizAcciones(1, 2)
               
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(9) VUELTA A TENDENCIA ALCISTA", "Gap Tasuki", MenComentario, MenConfirmacion, MenProcedimiento
         
            End If
            
         End If
      
      End If
    
   ' Si la vela de ayer y anteayer son negras y la de hoy es blanca
   ElseIf MatrizAcciones(3, 7) < -10 And MatrizAcciones(2, 7) < -10 And MatrizAcciones(1, 7) > 10 Then
   
      ' Si el maximo de la vela de ayer es menor al minimo de la vela de anteayer
      If MatrizAcciones(2, 4) < MatrizAcciones(3, 5) Then
      
         ' Si la apertura de hoy esta dentro del cuerpo de la vela de ayer
         If MatrizAcciones(1, 2) > MatrizAcciones(2, 3) And MatrizAcciones(1, 2) < MatrizAcciones(2, 2) Then
      
            ' Si el cierre de hoy esta en el hueco sin cerrarlo (por encima del maximo de ayer y por debajo del minimo de anteayer)
            If MatrizAcciones(1, 3) > MatrizAcciones(2, 4) And MatrizAcciones(1, 3) < MatrizAcciones(3, 5) Then
               
               MenComentario = "Hueco en velas negras de ayer y anteayer, con vela blanca hoy que abrio en el cuerpo de la vela de ayer y cierra en el hueco sin cerrarlo"
               MenConfirmacion = "Confirmación arriesgada por la formación en si, siempre que el heuco sea el primiero del movimiento bajista"
               MenProcedimiento = "Vender si se rompe mañana el mínimo de la primera hora o si el cierre de mañana es inferior a la vela de hoy " & MatrizAcciones(1, 3)
               
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(9) VUELTA A TENDENCIA BAJISTA", "Gap Tasuki", MenComentario, MenConfirmacion, MenProcedimiento
         
            End If
            
         End If
      
      End If
    
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_GemelosBlancos(Mercado As Integer)

 ' Tres velas las dos primeras son velas blancas entre las que hay un gap alcista
 ' la tercera blanca abriendo y cerrando más o menos al nivel de la segunda vela.
' Más significativo es si es el primero de un movimiento alcista reciente.

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If GemelosBlancos_Aviso = True Then
   
   ' Si la vela de ayer y anteayer son blancas y la de hoy es blanca
   If MatrizAcciones(3, 7) > 10 And MatrizAcciones(2, 7) > 10 And MatrizAcciones(1, 7) > 10 Then
   
      ' Si el minimo de la vela de ayer es mayor al maximo de la vela de anteayer
      If MatrizAcciones(2, 5) > MatrizAcciones(3, 4) Then
      
         ' Si la apertura de ayer menos la apertura de hoy es menor a un 0,2% de la apertura de ayer
         If Abs(MatrizAcciones(2, 2) - MatrizAcciones(1, 2)) * 500 < MatrizAcciones(2, 2) Then
            
            ' Si la cierre de ayer menos el cierre de hoy es menor a un 0,2% del cierre de ayer
            If Abs(MatrizAcciones(2, 3) - MatrizAcciones(1, 3)) * 500 < MatrizAcciones(2, 2) Then
               
               MenComentario = "Hueco alcista en velas blancas de ayer y anteayer, con vela blanca de hoy gemela a la de ayer"
               MenConfirmacion = "Solo es importante si el el primero de un movimiento alcista reciente"
               MenProcedimiento = ""
               
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(10) VUELTA A TENDENCIA ALCISTA", "Gemelos Blancos", MenComentario, MenConfirmacion, MenProcedimiento
         
            End If
               
         End If
      
      End If
    
   ' Si la vela de anteayer es negra y las de ayer y hoy blancas
   ElseIf MatrizAcciones(3, 7) < -10 And MatrizAcciones(2, 7) > 10 And MatrizAcciones(1, 7) > 10 Then
   
      ' Si el maximo de la vela de ayer es menor al minimo de la vela de anteayer
      If MatrizAcciones(2, 4) < MatrizAcciones(3, 5) Then
      
         ' Si la apertura de ayer menos la apertura de hoy es menor a un 0,2% de la apertura de ayer
         If Abs(MatrizAcciones(2, 2) - MatrizAcciones(1, 2)) * 500 < MatrizAcciones(2, 2) Then
            
            ' Si la cierre de ayer menos el cierre de hoy es menor a un 0,2% del cierre de ayer
            If Abs(MatrizAcciones(2, 3) - MatrizAcciones(1, 3)) * 500 < MatrizAcciones(2, 2) Then
               
               MenComentario = "Hueco bajista entre vela negra de anteayer y vela blanca de ayer, con vela blanca de hoy gemela a la de ayer"
               MenConfirmacion = "Solo es importante si el el primero de un movimiento bajista reciente"
               MenProcedimiento = ""
               
               ' Grabamos el aviso
               GuardarAvisos Mercado, "(10) VUELTA A TENDENCIA BAJISTA", "Gemelos Blancos", MenComentario, MenConfirmacion, MenProcedimiento
         
            End If
               
         End If
      
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_LineasSeparacion(Mercado As Integer)

' Dos velas la primera negra y la segunda blanca con gran cuerpo
' La apertura de la segunda vela se realiza en nivel de apertura de la primera vela.
' Preferentemente que las sobras de la vela blanca sean inexistentes o casi.


Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If LineasSeparacion_Aviso = True Then
   
   ' Si la vela de ayer es negra media o grande y la de hoy grande y blanca
   If MatrizAcciones(2, 7) < -20 And MatrizAcciones(1, 7) > 20 Then
   
      ' Si la apertura de hoy es igual a la apertura de ayer
      If MatrizAcciones(1, 2) = MatrizAcciones(2, 2) Then

          MenComentario = "Vela de ayer negra con apertura en el mismo precio de apertura que la vela blanca de hoy"
          MenConfirmacion = "Preferentemente las sobras de la vela blanca de hoy deben ser inexistentes o casi"
          MenProcedimiento = "Confirmación al día siguiente con cierre superior al de hoy"
               
          ' Grabamos el aviso
          GuardarAvisos Mercado, "(11) VUELTA A TENDENCIA ALCISTA", "Líneas de separación", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
    
   ' Si la vela de ayer es blanca media o grande y la de hoy grande y negra
   ElseIf MatrizAcciones(2, 7) > 20 And MatrizAcciones(1, 7) < -20 Then
   
      ' Si la apertura de hoy es igual a la apertura de ayer
      If MatrizAcciones(1, 2) = MatrizAcciones(2, 2) Then

          MenComentario = "Vela de ayer blanca con apertura en el mismo precio de apertura que la vela negra de hoy"
          MenConfirmacion = "Preferentemente las sobras de la vela negra de hoy deben ser inexistentes o casi"
          MenProcedimiento = "Confirmación al día siguiente con cierre inferior al de hoy"
               
          ' Grabamos el aviso
          GuardarAvisos Mercado, "(11) VUELTA A TENDENCIA BAJISTA", "Líneas de separación", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_LineasUnion(Mercado As Integer)

' Dos velas la primera negra con gran cuerpo y la segunda blanca
' donde el cierre de la segunda esta más o menos en el mismo nivel del cierre de la primera.
' Preferentemente que las sobras de la vela blanca sean inexistentes o casi.
' Confirmación al día siguiente con un cierre superior al cierre de la vela blanca.

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If LineasUnion_Aviso = True Then
   
   ' Si la vela de ayer es negra media o grande y la de hoy grande y blanca
   If MatrizAcciones(2, 7) < -20 And MatrizAcciones(1, 7) > 20 Then
   
      ' Si el cierre de hoy es igual al cierre de ayer
      If MatrizAcciones(1, 3) = MatrizAcciones(2, 3) Then

          MenComentario = "Vela de ayer negra con cierre en el mismo precio de cierre que la vela blanca de hoy"
          MenConfirmacion = "Preferentemente las sobras de la vela blanca de hoy deben ser inexistentes o casi"
          MenProcedimiento = "Confirmación al día siguiente con cierre superior al de la vela blanca de hoy"
               
          ' Grabamos el aviso
          GuardarAvisos Mercado, "(12) VUELTA A TENDENCIA ALCISTA", "Líneas de unión", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
    
   ' Si la vela de ayer es blanca media o grande y la de hoy grande y negra
   ElseIf MatrizAcciones(2, 7) > 20 And MatrizAcciones(1, 7) < -20 Then
   
      ' Si el cierre de hoy es igual al cierre de ayer
      If MatrizAcciones(1, 3) = MatrizAcciones(2, 3) Then

          MenComentario = "Vela de ayer blanca con cierre en el mismo precio que el cierre de la vela negra de hoy"
          MenConfirmacion = "Preferentemente las sobras de la vela negra de hoy deben ser inexistentes o casi"
          MenProcedimiento = "Confirmación al día siguiente con cierre inferior al de la vela negra de hoy"
               
          ' Grabamos el aviso
          GuardarAvisos Mercado, "(12) VUELTA A TENDENCIA BAJISTA", "Líneas de unión", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_Puntapie(Mercado As Integer)

' Dos velas la primera marubozu negro y la segunda marubozu blanco
' entre ambas hay un gap alcista.

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If Puntapie_Aviso = True Then
   
   ' Si la vela de ayer es marubozu negro medio o grande y la de hoy marubozu blanco medio o grande
   If (MatrizAcciones(2, 7) = -21 Or MatrizAcciones(2, 7) = -31) And (MatrizAcciones(1, 7) = 21 Or MatrizAcciones(1, 7) = 31) Then
   
      ' Si la apertura de ayer es menor a la apertura de hoy
      If MatrizAcciones(2, 2) < MatrizAcciones(1, 2) Then
      
         MenComentario = "Marubozu negro ayer y marubozu blanco hoy con gap ascendente entre ambos"
         MenConfirmacion = ""
         MenProcedimiento = ""
               
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(13) VUELTA A TENDENCIA ALCISTA", "Puntapie (Kicking)", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
    
   ' Si la vela de ayer es marubozu blanco medio o grande y la de hoy marubozu negro medio o grande
   ElseIf (MatrizAcciones(2, 7) = 21 Or MatrizAcciones(2, 7) = 31) And (MatrizAcciones(1, 7) = -21 Or MatrizAcciones(1, 7) = -31) Then
   
      ' Si la apertura de ayer es mayor a la apertura de hoy
      If MatrizAcciones(2, 2) > MatrizAcciones(1, 2) Then
      
         MenComentario = "Marubozu blanco ayer y marubozu negro hoy con gap descendente entre ambos"
         MenConfirmacion = ""
         MenProcedimiento = ""
               
         ' Grabamos el aviso
         GuardarAvisos Mercado, "(13) VUELTA A TENDENCIA BAJISTA", "Puntapie (Kicking)", MenComentario, MenConfirmacion, MenProcedimiento
      
      End If
    
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_TresRios(Mercado As Integer)
  
' Tres velas, la primera es negra grande
' la segunda es un martillo negro que marca un nuevo mínimo (20 días)
' la tercera es una peonza blanca, que su cuerpo se ha formado dentro de la sombra baja del martillo.

Dim MenComentario As String
Dim MenConfirmacion As String
Dim MenProcedimiento As String

If TresRios_Aviso = True Then
   
   ' Si la vela de anteayer es grande y negra, la de ayer es martillo negro y la de hoy es peonza
   If MatrizAcciones(3, 7) < -20 And MatrizAcciones(2, 7) = -11 And MatrizAcciones(1, 7) = 13 Then
   
      ' Si el cierre de ayer es mayor al cierre de hoy y el minimo de ayer es menor a la apertura de hoy
      ' la peonza se ha formado en la sombra inferior del martillo
      If MatrizAcciones(2, 3) > MatrizAcciones(1, 3) And MatrizAcciones(2, 5) < MatrizAcciones(1, 2) Then
      
         ' Si el minimo del martillo es menor o igual al soporte de 20 días
         If MatrizAcciones(2, 5) <= MatrizTemporal(Mercado, 11, 1) Then
         
            MenComentario = "Vela de anteayer grande y negra, vela de ayer martillo negro (en zona de soporte) y vela de hoy peonza blanca posicionada en sombra del martillo"
            MenConfirmacion = "Esperar cierre por encima del martillo de ayer " & MatrizAcciones(2, 2)
            MenProcedimiento = ""
               
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(14) VUELTA A TENDENCIA ALCISTA", "Suelo de los tres rios", MenComentario, MenConfirmacion, MenProcedimiento
      
         End If
         
      End If
      
   ' Si la vela de anteayer es grande y blanca, la de ayer es estrella fugaz blanca y la de hoy es peonza negra
   ElseIf MatrizAcciones(3, 7) > 20 And MatrizAcciones(2, 7) = 12 And MatrizAcciones(1, 7) = -13 Then
   
      ' Si el cierre de ayer es mayor al cierre de hoy y el minimo de ayer es menor a la apertura de hoy
      ' la peonza se ha formado en la sombra inferior del martillo
      If MatrizAcciones(2, 3) > MatrizAcciones(1, 3) And MatrizAcciones(2, 5) < MatrizAcciones(1, 2) Then
      
         ' Si el minimo del martillo es mayor o igual a la resistencia de 20 días
         If MatrizAcciones(2, 4) >= MatrizTemporal(Mercado, 11, 3) Then
         
            MenComentario = "Vela de anteayer grande y blanca, vela de ayer estrella fugaz blanca (en zona de resistencia) y vela de hoy peonza negra posicionada en sombra de la estrella"
            MenConfirmacion = "Esperar cierre por debajo de la estrella fugaz de ayer " & MatrizAcciones(2, 3)
            MenProcedimiento = ""
               
            ' Grabamos el aviso
            GuardarAvisos Mercado, "(14) VUELTA A TENDENCIA BAJISTA", "Techo de los tres rios", MenComentario, MenConfirmacion, MenProcedimiento
      
         End If
         
      End If
      
   End If
   
End If

End Sub

Public Sub AnalisisEstrategia_PausaHarami(Mercado As Integer)

' Una gran vela, seguida de una peonza en posicion harami
' El mínimo de la segunda vela debe ser superior al mínimo de la primera vela.
' El cuerpo de la primera vela engloba el cuerpo de la segunda.
' La tendencia es alcista.
' El mínimo de las cinco últimas velas (partiendo de la peonza en posición harami, esta misma incluida) es sucesivamente superior al mínimo de la vela anterior.
' La segunda vela es blanca.
' Procede comprar cuando la cotización de la acción rompe el máximo de la segunda o de la primera vela

Dim TipoHarami As String
Dim NumMinimos As Integer
Dim MenConfirmacion As String
Dim MenProcedimiento As String

TipoHarami = ""


If PausaHarami_Aviso = True Then

   ' Si la vela de hace antes de ayer es grande y blanca
   If MatrizAcciones(3, 7) > 20 Then
   
      
      ' Si la vela de antes de ayer contine en posición harami a la vela de ayer
      If MatrizAcciones(3, 3) > MatrizAcciones(2, 2) And MatrizAcciones(3, 3) > MatrizAcciones(2, 3) And MatrizAcciones(3, 2) < MatrizAcciones(2, 2) And MatrizAcciones(3, 2) < MatrizAcciones(2, 3) Then
        
         
         ' Si el minimo de ayer es mayor que el minimo de antes de ayer
         If MatrizAcciones(2, 5) > MatrizAcciones(3, 5) Then
         
        
            
            ' CAMBIADO Si la tendencia de corto es alcista o lateral alcista
            If InStr(1, MatrizTemporal(Mercado, 18, 8), "BAJISTA", vbTextCompare) = 0 Then
                
               ' Si el cierre de hoy es mayor al maximo de ayer
               If MatrizAcciones(1, 3) > MatrizAcciones(2, 4) Then
               
                  MenConfirmacion = "CONFIRMADO rotura de maximo de ayer " & MatrizAcciones(2, 4) & " "
                  MenProcedimiento = "Stop Loss en minimo vela de ayer " & MatrizAcciones(2, 5) & " y Objetivo revisar zonas de resistencia"
                  
                  TipoHarami = "A"
                  
               End If
               
               ' Si el cierre de hoy es mayor al maximo de antes de ayer
               If MatrizAcciones(1, 3) > MatrizAcciones(3, 4) Then
               
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de maximo de antes de ayer " & MatrizAcciones(3, 4) & " "
                  MenProcedimiento = "Stop Loss en minimo vela de ayer " & MatrizAcciones(2, 5) & " y Objetivo revisar zonas de resistencia"
                  
                  TipoHarami = "A"
                  
               End If
               
            End If
      
         End If
         
      End If
      
   ' Si la vela de hace antes de ayer es grande y negra
   ElseIf MatrizAcciones(3, 7) < -20 Then
   
      ' Si la vela de antes de ayer contine en posición harami a la vela de ayer
      If MatrizAcciones(3, 2) > MatrizAcciones(2, 2) And MatrizAcciones(3, 2) > MatrizAcciones(2, 3) And MatrizAcciones(3, 3) < MatrizAcciones(2, 2) And MatrizAcciones(3, 3) < MatrizAcciones(2, 3) Then
      
         ' Si el maximo de ayer es menor que el maximo de antes de ayer
         If MatrizAcciones(2, 4) < MatrizAcciones(3, 4) Then
         
            ' Si la tendencia de medio es bajista o lateral bajista
            If InStr(1, MatrizTemporal(Mercado, 16, 8), "BAJISTA", vbTextCompare) <> 0 Then
            
               ' Si el cierre de hoy es menor al minimo de ayer
               If MatrizAcciones(1, 3) < MatrizAcciones(2, 5) Then
               
                  MenConfirmacion = "CONFIRMADO rotura de minimo de ayer " & MatrizAcciones(2, 5) & " "
                  MenProcedimiento = "Stop Loss en maximo vela de ayer " & MatrizAcciones(2, 4) & " y Objetivo revisar zonas de soporte"
                  
                  TipoHarami = "B"
                  
               End If
               
               ' Si el cierre de hoy es menor al minimo de antes de ayer
               If MatrizAcciones(1, 3) < MatrizAcciones(3, 5) Then
               
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de minimo de antes de ayer " & MatrizAcciones(3, 5) & " "
                  MenProcedimiento = "Stop Loss en maximo vela de ayer " & MatrizAcciones(2, 4) & " y Objetivo revisar zonas de soporte"
                  
                  TipoHarami = "B"
                  
               End If
               
            End If
      
         End If
      
      End If
   
   End If
   
   ' Si existe pausa harami
   If TipoHarami <> "" Then
      
      NumMinimos = 0
      
      ' Si es alcista
      If TipoHarami = "A" Then
         
         ' Recorremos desde ayer hasta 5 días más atras
         For Y = 2 To 5
             
             ' Si el minimo tratado es mayor al minimo del dia anterior
             If MatrizAcciones(Y, 5) > MatrizAcciones(Y + 1, 5) Then
                
                NumMinimos = NumMinimos + 1
             
             End If
                      
         Next
         
         ' Si el numero de minimos consecutivamente superiores son 5
         If NumMinimos >= 5 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cinco minimos superiores hasta la peonza "
            
         End If
         
         ' Si la segunda vela es blanca
         If MatrizAcciones(2, 7) > 0 Then MenConfirmacion = MenConfirmacion & "CONFIRMADO peonza blanca"
         
         ' Si la segunda vela es una peonza o doji
         If Abs(MatrizAcciones(2, 7)) < 20 Then
            
             MenConfirmacion = "CONFIRMADO vela de ayer es peonza " & MenConfirmacion
             
         ' Si no lo es
         Else
         
            MenConfirmacion = "OJO VELA DE AYER NO ES PEONZA " & MenConfirmacion
         
         End If
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "COMPRA", "Pausa Harami", "Pausa Harami alcista, porque vela de antes de ayer contiene a vela de ayer y hoy se ha roto el maximo de ayer o antes de ayer", MenConfirmacion, MenProcedimiento
         
         
      ' Si es bajista
      Else
         
         ' Recorremos desde ayer hasta 5 días más atras
         For Y = 2 To 5
             
             ' Si el maximo tratado es menor al maximo del dia anterior
             If MatrizAcciones(Y, 4) < MatrizAcciones(Y + 1, 4) Then
                
                NumMinimos = NumMinimos + 1
             
             End If
                      
         Next
         
         ' Si el numero de minimos consecutivamente superiores son 5
         If NumMinimos >= 5 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cinco maximos inferiores hasta la peonza "
            
         End If
         
         ' Si la segunda vela es negra
         If MatrizAcciones(2, 7) < 0 Then MenConfirmacion = MenConfirmacion & "CONFIRMADO peonza negra"
         
         ' Si la segunda vela es una peonza o doji
         If Abs(MatrizAcciones(2, 7)) < 20 Then
            
             MenConfirmacion = "CONFIRMADO vela de ayer es peonza " & MenConfirmacion
             
         ' Si no lo es
         Else
         
            MenConfirmacion = "OJO VELA DE AYER NO ES PEONZA " & MenConfirmacion
         
         End If
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "VENTA", "Pausa Harami", "Pausa Harami  bajista, porque vela de antes de ayer contiene a vela de ayer y hoy se ha roto el minimo de ayer o antes de ayer", MenConfirmacion, MenProcedimiento
         
      End If
      
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_PausaCasiHarami(Mercado As Integer)

' Una gran vela, seguida de una peonza en posicion casi harami (el cuerpo se desborda hacia una de las sombras)
' El mínimo de la segunda vela debe ser superior al mínimo de la primera vela.
' El cuerpo de la primera vela engloba el cuerpo de la segunda.
' La tendencia es alcista.
' El mínimo de las cinco últimas velas (partiendo de la peonza en posición harami, esta misma incluida) es sucesivamente superior al mínimo de la vela anterior.
' La segunda vela es blanca.
' Procede comprar cuando la cotización de la acción rompe el máximo de la segunda o de la primera vela

Dim TipoHarami As String
Dim NumMinimos As Integer
Dim MenConfirmacion As String
Dim MenProcedimiento As String

TipoHarami = ""

If PausaCasiHarami_Aviso = True Then

   ' Si la vela de hace antes de ayer es grande y blanca
   If MatrizAcciones(3, 7) > 20 Then
   
      ' la vela de ayer esta en casi harami, desbordada hacia el maximo de la vela precedente
      If MatrizAcciones(3, 4) > MatrizAcciones(2, 2) And MatrizAcciones(3, 4) > MatrizAcciones(2, 3) And (MatrizAcciones(3, 3) < MatrizAcciones(2, 2) Or MatrizAcciones(3, 3) < MatrizAcciones(2, 3)) Then
      
         ' Si el minimo de ayer es mayor que el minimo de antes de ayer
         If MatrizAcciones(2, 5) > MatrizAcciones(3, 5) Then
         
            ' Si la tendencia de medio es alcista o lateral alcista
            If InStr(1, MatrizTemporal(Mercado, 16, 8), "ALCISTA", vbTextCompare) <> 0 Then
                
               ' Si el cierre de hoy es mayor al maximo de ayer
               If MatrizAcciones(1, 3) > MatrizAcciones(2, 4) Then
               
                  MenConfirmacion = "CONFIRMADO rotura de maximo de ayer " & MatrizAcciones(2, 4) & " "
                  MenProcedimiento = "Stop Loss en minimo vela de ayer " & MatrizAcciones(2, 5) & " y Objetivo revisar zonas de resistencia"
                  
                  TipoHarami = "A"
                  
               End If
               
               ' Si el cierre de hoy es mayor al maximo de antes de ayer
               If MatrizAcciones(1, 3) > MatrizAcciones(3, 4) Then
               
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de maximo de antes de ayer " & MatrizAcciones(3, 4) & " "
                  MenProcedimiento = "Stop Loss en minimo vela de ayer " & MatrizAcciones(2, 5) & " y Objetivo revisar zonas de resistencia"
                  
                  TipoHarami = "A"
                  
               End If
               
            End If
      
         End If
         
      End If
      
   ' Si la vela de hace antes de ayer es grande y negra
   ElseIf MatrizAcciones(3, 7) < -20 Then
   
      ' Si la vela de antes de ayer contine en posición harami a la vela de ayer
      If MatrizAcciones(3, 5) < MatrizAcciones(2, 2) And MatrizAcciones(3, 5) < MatrizAcciones(2, 3) And (MatrizAcciones(3, 3) < MatrizAcciones(2, 2) Or MatrizAcciones(3, 3) < MatrizAcciones(2, 3)) Then
            
      
         ' Si el maximo de ayer es menor que el maximo de antes de ayer
         If MatrizAcciones(2, 4) < MatrizAcciones(3, 4) Then
         
            ' Si la tendencia de medio es bajista o lateral bajista
            If InStr(1, MatrizTemporal(Mercado, 16, 8), "BAJISTA", vbTextCompare) <> 0 Then
            
               ' Si el cierre de hoy es menor al minimo de ayer
               If MatrizAcciones(1, 3) < MatrizAcciones(2, 5) Then
               
                  MenConfirmacion = "CONFIRMADO rotura de minimo de ayer " & MatrizAcciones(2, 5) & " "
                  MenProcedimiento = "Stop Loss en maximo vela de ayer " & MatrizAcciones(2, 4) & " y Objetivo revisar zonas de soporte"
                  
                  TipoHarami = "B"
                  
               End If
               
               ' Si el cierre de hoy es menor al minimo de antes de ayer
               If MatrizAcciones(1, 3) < MatrizAcciones(3, 5) Then
               
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de minimo de antes de ayer " & MatrizAcciones(3, 5) & " "
                  MenProcedimiento = "Stop Loss en maximo vela de ayer " & MatrizAcciones(2, 4) & " y Objetivo revisar zonas de soporte"
                  
                  TipoHarami = "B"
                  
               End If
               
            End If
      
         End If
      
      End If
   
   End If
   
   ' Si existe pausa harami
   If TipoHarami <> "" Then
      
      NumMinimos = 0
      
      ' Si es alcista
      If TipoHarami = "A" Then
         
         ' Recorremos desde ayer hasta 5 días más atras
         For Y = 2 To 5
             
             ' Si el minimo tratado es mayor al minimo del dia anterior
             If MatrizAcciones(Y, 5) > MatrizAcciones(Y + 1, 5) Then
                
                NumMinimos = NumMinimos + 1
             
             End If
                      
         Next
         
         ' Si el numero de minimos consecutivamente superiores son 5
         If NumMinimos >= 5 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cinco minimos superiores hasta la peonza "
            
         End If
         
         ' Si la segunda vela es blanca
         If MatrizAcciones(2, 7) > 0 Then MenConfirmacion = MenConfirmacion & "CONFIRMADO peonza blanca"
         
         ' Si la segunda vela es una peonza o doji
         If Abs(MatrizAcciones(2, 7)) < 20 Then
            
             MenConfirmacion = "CONFIRMADO vela de ayer es peonza " & MenConfirmacion
             
         ' Si no lo es
         Else
         
            MenConfirmacion = "OJO VELA DE AYER NO ES PEONZA " & MenConfirmacion
         
         End If
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "COMPRA", "Pausa Casi Harami", "Pausa Casi Harami alcista, porque vela de antes de ayer contiene a vela de ayer desbordada al maximo y hoy se ha roto el maximo de ayer o antes de ayer", MenConfirmacion, MenProcedimiento
         
         
      ' Si es bajista
      Else
         
         ' Recorremos desde ayer hasta 5 días más atras
         For Y = 2 To 5
             
             ' Si el maximo tratado es menor al maximo del dia anterior
             If MatrizAcciones(Y, 4) < MatrizAcciones(Y + 1, 4) Then
                
                NumMinimos = NumMinimos + 1
             
             End If
                      
         Next
         
         ' Si el numero de minimos consecutivamente superiores son 5
         If NumMinimos >= 5 Then
         
            MenConfirmacion = MenConfirmacion & "CONFIRMADO cinco maximos inferiores hasta la peonza "
            
         End If
         
         ' Si la segunda vela es negra
         If MatrizAcciones(2, 7) < 0 Then MenConfirmacion = MenConfirmacion & "CONFIRMADO peonza negra"
         
         ' Si la segunda vela es una peonza o doji
         If Abs(MatrizAcciones(2, 7)) < 20 Then
            
             MenConfirmacion = "CONFIRMADO vela de ayer es peonza " & MenConfirmacion
             
         ' Si no lo es
         Else
         
            MenConfirmacion = "OJO VELA DE AYER NO ES PEONZA " & MenConfirmacion
         
         End If
         
         ' Grabamos el aviso
         GuardarAvisos Mercado, "VENTA", "Pausa Casi Harami", "Pausa Casi Harami  bajista, porque vela de antes de ayer contiene a vela de ayer desbordada hacia el minimo y hoy se ha roto el minimo de ayer o antes de ayer", MenConfirmacion, MenProcedimiento
         
      End If
      
   End If
    
End If

End Sub

Public Sub AnalisisEstrategia_CorreccionOrdenada(Mercado As Integer)

' Se compone de 3 a 5 velas negras
' Al menos tres de las velas tiene un maximo inferior al maximo precedente
' Tambien es recomendable que el minimo sea inferior al minimo precedente
' Se alcanzo un nuevo maximo en menos de ocho barras anteriores a la de entrada
' Aparece un intento de cambio de tendencia en la ultima vela de la correcion
' La vela de entrada o la ultima de la correción forma un martillo
' Existe un soporte en forma de línea de tendencia o de media móvil
' La media movil de 20 o de 50 apuntaba haca arriba en el momento que la correción aparecio

Dim TipoCorreccion As String
Dim NumCorreccion As Integer
Dim NumVelas As Integer
Dim ReglaVelas As Integer
Dim ReglaVelas2 As Integer
Dim MaxAnterior As Integer
Dim CambioTendencia As Boolean
Dim VelaMartillo As Boolean
Dim SopTendencia As Boolean
Dim MM20Positiva As Boolean
Dim MM50Positiva As Boolean
Dim MenConfirmacion As String
Dim MenProcedimiento As String

TipoCorreccion = ""
NumCorreccion = 0
NumVelas = 0
ReglaVelas = 0
ReglaVelas2 = 0
MaxAnterior = 0
CambioTendencia = False 'De momento no lo utilizamos y lo avisamos
VelaMartillo = False
SopTendencia = False 'De momento no lo utilizamos y lo avisamos
MM20Positiva = False
MM50Positiva = False

If CorreccionOrdenada_Aviso = True Then
   
   ' Si la Vela actual es negra
   If MatrizAcciones(1, 7) < 0 Then
   
      ' Si la vela de ayer es negra
      If MatrizAcciones(2, 7) < 0 Then
      
         TipoCorreccion = "B"
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(1, 7)) = 11 Then VelaMartillo = True
      
      ' Si la vela de ayer es blanca
      Else
      
         TipoCorreccion = "A"
         NumCorreccion = 1
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(1, 7)) = 11 Then VelaMartillo = True
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(2, 7)) = 11 Then VelaMartillo = True
      
      End If
   
   ' Si la Vela actual es blanca
   Else
   
      ' Si la vela de ayer es negra
      If MatrizAcciones(2, 7) < 0 Then
      
         TipoCorreccion = "B"
         NumCorreccion = 1
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(1, 7)) = 11 Then VelaMartillo = True
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(2, 7)) = 11 Then VelaMartillo = True
      
      ' Si la vela de ayer es blanca
      Else
      
         TipoCorreccion = "A"
         
         ' Si la vela es un martillo
         If Abs(MatrizAcciones(1, 7)) = 11 Then VelaMartillo = True
      
      End If
   
   End If
   
   ' Si hay previsión de correcion
   If TipoCorreccion <> "" Then
   
      ' Recorremos velas desde hoy si no es confirmada o ayer si es confirmada
      For Y = 1 + NumCorreccion To 5 + NumCorreccion
          
          ' Si la correcion es bajista
          If TipoCorrecion = "B" Then
          
             ' Si el MAX de hoy tratado es menor que el MAX del dia anterior
             If MatrizAcciones(Y, 4) < MatrizAcciones(Y + 1, 4) Then
             
                ReglaVelas = ReglaVelas + 1
             
             End If
             
             ' Si el MIN de hoy tratado es menor que el MIN del dia anterior
             If MatrizAcciones(Y, 5) < MatrizAcciones(Y + 1, 5) Then
             
                ReglaVelas2 = ReglaVelas2 + 1
             
             End If
          
             ' Si la vela es negra
             If MatrizAcciones(Y, 7) < 0 Then
                
                NumVelas = NumVelas + 1
                
             ' Si la vela es blanca y esta dentro de las tres ultimas
             ElseIf Y < (1 + NumCorreccion + 3) Then
             
                NumVelas = 0
                
                Y = 5 + NumCorreccion
             
             End If
          
          ' Si la correcion es alcista
          Else
          
             ' Si el MIN de hoy tratado es mayor que el MIN del dia anterior
             If MatrizAcciones(Y, 5) > MatrizAcciones(Y + 1, 5) Then
             
                ReglaVelas = ReglaVelas + 1
             
             End If
             
             ' Si el MAX de hoy tratado es mayor que el MAX del dia anterior
             If MatrizAcciones(Y, 4) > MatrizAcciones(Y + 1, 4) Then
             
                ReglaVelas2 = ReglaVelas2 + 1
             
             End If
            
             ' Si la vela es blanca
             If MatrizAcciones(Y, 7) > 0 Then
             
                NumVelas = NumVelas + 1
                
             ' Si la vela es negra y esta dentro de las tres ultimas
             ElseIf Y < (1 + NumCorreccion + 3) Then
             
                NumVelas = 0
                
                Y = 5 + NumCorreccion
             
             End If
          
          End If
   
      Next
   
   End If
   
   ' Si la correcion es de 3 o más velas, y se cumple la regla de minimos inferiores o maximos superiores
   If NumVelas >= 3 And ReglaVelas >= 3 Then
   
      ' Si la correcion es bajista
      If TipoCorrecion = "B" Then
      
         ' Si la MM20 anterior a la correción es menor a la MM20 de la primera vela de la correcion
         If MatrizTemporal(Mercado, 2, NumVelas + NumCorreccion + 2) < MatrizTemporal(Mercado, 2, NumVelas + NumCorreccion + 1) Then
         
            MM20Positiva = True
         
         End If
         
         ' Si la MM50 anterior a la correción es menor a la MM50 de la primera vela de la correcion
         If MatrizTemporal(Mercado, 3, NumVelas + NumCorreccion + 2) < MatrizTemporal(Mercado, 3, NumVelas + NumCorreccion + 1) Then
         
            MM50Positiva = True
         
         End If
         
      
      ' Si la correcion es alcista
      Else
      
         ' Si la MM20 anterior a la correción es mayor a la MM20 de la primera vela de la correcion
         If MatrizTemporal(Mercado, 2, NumVelas + NumCorreccion + 2) > MatrizTemporal(Mercado, 2, NumVelas + NumCorreccion + 1) Then
         
            MM20Positiva = True
         
         End If
         
         ' Si la MM50 anterior a la correción es menor a la MM50 de la primera vela de la correcion
         If MatrizTemporal(Mercado, 3, NumVelas + NumCorreccion + 2) > MatrizTemporal(Mercado, 3, NumVelas + NumCorreccion + 1) Then
         
            MM50Positiva = True
         
         End If
      
      End If
   
      MaxAnterior = 10
   
      ' Recorremos velas desde hoy si no es confirmada o ayer si es confirmada
      For Y = 10 To 1 + NumCorreccion Step -1
   
          ' Si la correcion es bajista
          If TipoCorrecion = "B" Then
          
             ' Si el maximo de la cotización tratada es mayor que el anterior
             If MatrizAcciones(Y, 4) > MatrizAcciones(MaxAnterior, 4) Then
             
                MaxAnterior = Y
             
             End If
          
          Else
             
             ' Si el minimo de la cotización tratada es menor que el anterior
             If MatrizAcciones(Y, 5) < MatrizAcciones(MaxAnterior, 5) Then
             
                MaxAnterior = Y
             
             End If
          
          End If
      
      Next
      
      ' Si el maximo para bajista o el minimo para alcista se produjo en menos de ocho velas
      If MaxAnterior < 8 Then
      
         ' Si la correcion es bajista
         If TipoCorrecion = "B" Then
         
            MenConfirmacion = ""
            
            If NumCorreccion = 1 Then
            
               ' Si el maximo de hoy supero el maximo de ayer
               If MatrizAcciones(1, 4) > MatrizAcciones(2, 4) Then
            
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de max última vela " & CStr(MatrizAcciones(2, 4)) & " "
               
                  MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 5) & " o en " & MatrizAcciones(2, 5) & " Objetivo " & MatrizAcciones(MaxAnterior, 4)
                  
               Else
               
                  MenConfirmacion = MenConfirmacion & "Confirmar rotura de max última vela " & CStr(MatrizAcciones(2, 4)) & " "
                  
                  MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 5) & " Objetivo " & MatrizAcciones(MaxAnterior, 4)
                  
               End If
               
            Else
            
               MenConfirmacion = MenConfirmacion & "Confirmar rotura de max última vela " & CStr(MatrizAcciones(1, 4)) & " "
               
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 5) & " Objetivo " & MatrizAcciones(MaxAnterior, 4)
               
            End If
            
            If VelaMartillo = True Then MenConfirmacion = MenConfirmacion & "Martillo en confirmación o en ultima corrección "
            
            If MM20Positiva = True Then MenConfirmacion = MenConfirmacion & "MM20 alcista en inicio "
            
            If MM50Positiva = True Then MenConfirmacion = MenConfirmacion & "MM50 alcista en inicio "
      
            MenConfirmacion = MenConfirmacion & "Confirmar cambio tendencia o soporte en ultima vela"
            
            'MsgBox "COMPRA Correccion ordenada"
            
            GuardarAvisos Mercado, "COMPRA", "Correccion Ordenada", "Compuesta por " & CStr(NumVelas) & " velas, con " & CStr(ReglaVelas) & " maximos inferiores y " & CStr(ReglaVelas2) & " minimos inferiores, maximo alcanzado en " & CStr(MaxAnterior), MenConfirmacion, MenProcedimiento
            
         ' Si la correcion es alcista
         Else
         
            MenConfirmacion = ""
            
            If NumCorreccion = 1 Then
            
               ' Si el minimo de hoy revaso a la baja el minimo de ayer
               If MatrizAcciones(1, 5) < MatrizAcciones(2, 5) Then
            
                  MenConfirmacion = MenConfirmacion & "CONFIRMADO rotura de min última vela " & CStr(MatrizAcciones(2, 5)) & " "
               
                  MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 4) & " o en " & MatrizAcciones(2, 4) & " Objetivo " & MatrizAcciones(MaxAnterior, 5)
               
               Else
               
                  MenConfirmacion = MenConfirmacion & "Confirmar rotura de min última vela " & CStr(MatrizAcciones(2, 5)) & " "
                  
                  MenProcedimiento = "Stop Loss en " & MatrizAcciones(2, 4) & " Objetivo " & MatrizAcciones(MaxAnterior, 5)
                  
               End If
               
            Else
            
               MenConfirmacion = MenConfirmacion & "Confirmar rotura de min última vela " & CStr(MatrizAcciones(1, 5)) & " "
               
               MenProcedimiento = "Stop Loss en " & MatrizAcciones(1, 4) & " Objetivo " & MatrizAcciones(MaxAnterior, 5)
               
            End If
            
            If VelaMartillo = True Then MenConfirmacion = MenConfirmacion & "Martillo en confirmación o en ultima corrección "
            
            If MM20Positiva = True Then MenConfirmacion = MenConfirmacion & "MM20 bajista en inicio "
            
            If MM50Positiva = True Then MenConfirmacion = MenConfirmacion & "MM50 bajista en inicio "
      
            MenConfirmacion = MenConfirmacion & "Confirmar cambio tendencia o resistencia en ultima vela"
            
            'MsgBox "VENTA Correccion ordenada"
            
            GuardarAvisos Mercado, "VENTA", "Correccion Ordenada", "Compuesta por " & CStr(NumVelas) & " velas, con " & CStr(ReglaVelas) & " minimos superiores y " & CStr(ReglaVelas2) & " maximos superiores, minimo alcanzado en " & CStr(MaxAnterior), MenConfirmacion, MenProcedimiento
        
         End If
      
      End If
       
   End If

End If

End Sub

Public Sub AnalisisEstrategia_Lane(Mercado As Integer)

' CALCULO DEL PUNTO DE CRUCE ENTRE k y d
' Interseccion = ((kn1 * dn) - (kn * dn1)) / ((kn1 - kn) - (dn1 - dn))

Dim Interseccion As Double

' CLASICA LANE LENTO

' Si DS corta de abajo a arriba a DSS en zona de sobreventa COMPRA
' Si DS corta de arriba a abajo a DSS en zona de sobrecompra VENTA
' Confirmación despues del corte, diferencia de 4 o 5 con DSS
' Sobreconfirmación en compra la siguiente sesión debe superar el maximo de la anterior
' Sobreconfirmación en venta la siguiente sesión debe superar el minimo a la baja de la anterior

If Lane_AvisoClasica_Lento = True Then
       
   ' Si el valor actual de K es mayor que D
   If MatrizTemporal(Mercado, 22, 1) > MatrizTemporal(Mercado, 23, 1) Then
      
      ' Recorremos valores de ayer hasta n-4
      For Y = 2 To 3
          
          ' Si el valor actual de K es menor que D
          If MatrizTemporal(Mercado, 22, Y) < MatrizTemporal(Mercado, 23, Y) Then

             Interseccion = ((MatrizTemporal(Mercado, 22, Y - 1) * MatrizTemporal(Mercado, 23, Y)) - (MatrizTemporal(Mercado, 22, Y) * MatrizTemporal(Mercado, 23, Y - 1))) / ((MatrizTemporal(Mercado, 22, Y - 1) - MatrizTemporal(Mercado, 22, Y)) - (MatrizTemporal(Mercado, 23, Y - 1) - MatrizTemporal(Mercado, 23, Y)))
             
             'MsgBox "Interseccion: " & Interseccion
             
             ' Si el punto de cruce entre k y d es menor o igual a la zona de sobreventa
             If Interseccion <= Lane_SobreVenta Then
                             
                ' Si K actual es mayor que D más 4 (límite de confirmación)
                If MatrizTemporal(Mercado, 22, 1) > (MatrizTemporal(Mercado, 23, 1) + 4) Then
                   
                   ' Si el maximo de la cotización actual es mayor al maximo de la cotización del cruce
                   If MatrizAcciones(1, 4) > MatrizAcciones(Y, 4) Then
                   
                      GuardarAvisos Mercado, "COMPRA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "confirmación de que K es mayor que D + 4 y sobreconfirmación maximo mayor que maximo del cruce", ""
                   
                   Else
                      
                      GuardarAvisos Mercado, "COMPRA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "confirmación de que K es mayor que D + 4", ""
                                      
                   End If
                
                Else
                
                   GuardarAvisos Mercado, "COMPRA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "esperar confirmación (K > (D + 4)) y sobreconfirmación Maximo cotización superior a Maximo cotización cruce", ""
                   
                   
                End If

             End If
             
             ' Forzamos la salida de For
             Y = 5
             
          End If
      
      Next
      
   ' Si el valor actual de K es menor que D
   ElseIf MatrizTemporal(Mercado, 22, 1) < MatrizTemporal(Mercado, 23, 1) Then
      
      ' Recorremos valores de ayer hasta n-4
      For Y = 2 To 3
          
          ' Si el valor actual de K es mayor que D
          If MatrizTemporal(Mercado, 22, Y) > MatrizTemporal(Mercado, 23, Y) Then
             
             Interseccion = ((MatrizTemporal(Mercado, 22, Y - 1) * MatrizTemporal(Mercado, 23, Y)) - (MatrizTemporal(Mercado, 22, Y) * MatrizTemporal(Mercado, 23, Y - 1))) / ((MatrizTemporal(Mercado, 22, Y - 1) - MatrizTemporal(Mercado, 22, Y)) - (MatrizTemporal(Mercado, 23, Y - 1) - MatrizTemporal(Mercado, 23, Y)))
             
             'MsgBox "Interseccion: " & Interseccion
             
             ' Si el punto de cruce entre k y d es mayor o igual a la zona de sobrecompra
             If Interseccion >= Lane_SobreCompra Then
                                 
                ' Si K actual más 4 (límite de confirmación) es menor que D
                If (MatrizTemporal(Mercado, 22, 1) + 4) < MatrizTemporal(Mercado, 23, 1) Then
                   
                   ' Si el minimo de la cotización actual es menor al minimo de la cotización del cruce
                   If MatrizAcciones(1, 5) < MatrizAcciones(Y, 5) Then
                   
                      GuardarAvisos Mercado, "VENTA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de arriba a abajo, en zona de sobrecompra (" & Lane_SobreCompra & ")", "confirmación de que K + 4 es menor que D y sobreconfirmación minimo menor que minimo del cruce", ""
                   
                   Else
                   
                      GuardarAvisos Mercado, "VENTA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobrecompra (" & Lane_SobreCompra & ")", "confirmación de que K + 4 es menor que D", ""

                   End If
                
                Else
                
                   GuardarAvisos Mercado, "VENTA", "Lane Lento Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobrecompra (" & Lane_SobreCompra & ")", "esperar confirmación (K + 4 < D) y sobreconfirmación minimo cotización inferior a minimo cotización cruce", ""

                End If
                
             End If
             
             ' Forzamos la salida de For
             Y = 5
             
          End If
      
      Next
       
   End If

End If

' CLASICA LANE RAPIDO

' Si K corta de abajo a arriba a D en zona de sobreventa COMPRA
' Si K corta de arriba a abajo a D en zona de sobrecompra VENTA
' Confirmación despues del corte, diferencia de 4 o 5 con D
' Sobreconfirmación en compra la siguiente sesión debe superar el maximo de la anterior
' Sobreconfirmación en venta la siguiente sesión debe superar el minimo a la baja de la anterior

If Lane_AvisoClasica_Rapido = True Then

   ' Si el valor actual de K es mayor que D
   If MatrizTemporal(Mercado, 20, 1) > MatrizTemporal(Mercado, 21, 1) Then
      
      ' Recorremos valores de ayer hasta n-4
      For Y = 2 To 3
          
          ' Si el valor actual de K es menor que D
          If MatrizTemporal(Mercado, 20, Y) < MatrizTemporal(Mercado, 21, Y) Then

             Interseccion = ((MatrizTemporal(Mercado, 20, Y - 1) * MatrizTemporal(Mercado, 21, Y)) - (MatrizTemporal(Mercado, 20, Y) * MatrizTemporal(Mercado, 21, Y - 1))) / ((MatrizTemporal(Mercado, 20, Y - 1) - MatrizTemporal(Mercado, 20, Y)) - (MatrizTemporal(Mercado, 21, Y - 1) - MatrizTemporal(Mercado, 21, Y)))
             
             'MsgBox "Interseccion: " & Interseccion
             
             ' Si el punto de cruce entre k y d es menor o igual a la zona de sobreventa
             If Interseccion <= Lane_SobreVenta Then
                
                ' Si K actual es mayor que D más 4 (límite de confirmación)
                If MatrizTemporal(Mercado, 20, 1) > (MatrizTemporal(Mercado, 21, 1) + 4) Then
                   
                   ' Si el maximo de la cotización actual es mayor al maximo de la cotización del cruce
                   If MatrizAcciones(1, 4) > MatrizAcciones(Y, 4) Then
                   
                      GuardarAvisos Mercado, "COMPRA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "confirmación de que K es mayor que D + 4 y sobreconfirmación maximo mayor que maximo del cruce", ""
                   
                   Else
                   
                      GuardarAvisos Mercado, "COMPRA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "confirmación de que K es mayor que D + 4", ""
                      
                   End If
                
                Else
                   
                   GuardarAvisos Mercado, "COMPRA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobreventa (" & Lane_SobreVenta & ")", "esperar confirmación (K > (D + 4)) y sobreconfirmación Maximo cotización superior a Maximo cotización cruce", ""
                   
                End If

             End If
             
             ' Forzamos la salida de For
             Y = 5
             
          End If
      
      Next
      
   ' Si el valor actual de K es menor que D
   ElseIf MatrizTemporal(Mercado, 20, 1) < MatrizTemporal(Mercado, 21, 1) Then
      
      ' Recorremos valores de ayer hasta n-4
      For Y = 2 To 3
          
          ' Si el valor actual de K es mayor que D
          If MatrizTemporal(Mercado, 20, Y) > MatrizTemporal(Mercado, 21, Y) Then
             
             Interseccion = ((MatrizTemporal(Mercado, 20, Y - 1) * MatrizTemporal(Mercado, 21, Y)) - (MatrizTemporal(Mercado, 20, Y) * MatrizTemporal(Mercado, 21, Y - 1))) / ((MatrizTemporal(Mercado, 20, Y - 1) - MatrizTemporal(Mercado, 20, Y)) - (MatrizTemporal(Mercado, 21, Y - 1) - MatrizTemporal(Mercado, 21, Y)))
                          
             ' Si el punto de cruce entre k y d es mayor o igual a la zona de sobrecompra
             If Interseccion >= Lane_SobreCompra Then
                 
                ' Si K actual más 4 (límite de confirmación) es menor que D
                If (MatrizTemporal(Mercado, 20, 1) + 4) < MatrizTemporal(Mercado, 21, 1) Then
                   
                   ' Si el minimo de la cotización actual es menor al minimo de la cotización del cruce
                   If MatrizAcciones(1, 5) < MatrizAcciones(Y, 5) Then
                   
                      GuardarAvisos Mercado, "VENTA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de arriba a abajo, en zona de sobrecompra (" & Lane_SobreCompra & ")", "confirmación de que K + 4 es menor que D y sobreconfirmación minimo menor que minimo del cruce", ""
                   
                   Else
                      
                      GuardarAvisos Mercado, "VENTA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobrecompra (" & Lane_SobreCompra & ")", "confirmación de que K + 4 es menor que D", ""
                      
                   End If
                
                Else
                   
                   GuardarAvisos Mercado, "VENTA", "Lane Rapido Estrategia Clasica", "Cruce de K con D (" & CStr(Round(Interseccion, 2)) & ") de abajo a arriba, en zona de sobrecompra (" & Lane_SobreCompra & ")", "esperar confirmación (K + 4 < D) y sobreconfirmación minimo cotización inferior a minimo cotización cruce", ""
                   
                End If
                
             End If
             
             ' Forzamos la salida de For
             Y = 5
             
          End If
      
      Next
       
   End If

End If

' SALIDA DE ZONA LANE LENTO

' Si DS sale de zona de sobreventa COMPRA
' Si DS sale de zona de sobrecompra VENTA
' El stop (NO CONTROLADO), se pone cuando DS vuelve a la zona de la que salio

If Lane_AvisoSZona_Lento = True Then

   ' Si el valor de DSn > que DSn-1 probamos compra saliendo de zona de sobreventa
   If MatrizTemporal(Mercado, 22, 1) > MatrizTemporal(Mercado, 22, 2) Then
   
      ' Si DSn >= que la zona de Sobreventa y DSn-1 era menor que la zona de SobreVenta
      If MatrizTemporal(Mercado, 22, 1) >= Lane_SobreVenta And MatrizTemporal(Mercado, 22, 2) < Lane_SobreVenta Then
      
         GuardarAvisos Mercado, "COMPRA", "Lane Lento Estrategia Salida de Zona", "Salida de DS (" & CStr(Round(MatrizTemporal(Mercado, 22, 1), 2)) & ") de zona de sobreventa (" & Lane_SobreVenta & ")", "", ""
      
      End If
   
   ' Si el valor de DSn < que DSn-1 probamos compra saliendo de zona de sobrecompra
   ElseIf MatrizTemporal(Mercado, 22, 1) < MatrizTemporal(Mercado, 22, 2) Then
   
      ' Si DSn <= que la zona de Sobrecompra y DSn-1 era mayor que la zona de SobreCompra
      If MatrizTemporal(Mercado, 22, 1) <= Lane_SobreCompra And MatrizTemporal(Mercado, 22, 2) > Lane_SobreCompra Then
      
         GuardarAvisos Mercado, "VENTA", "Lane Lento Estrategia Salida de Zona", "Salida de DS (" & CStr(Round(MatrizTemporal(Mercado, 22, 1), 2)) & ") de zona de sobrecompra (" & Lane_SobreCompra & ")", "", ""
      
      End If
   
   End If
   
End If

' SALIDA DE ZONA LANE RAPIDO

' Si D sale de zona de sobreventa COMPRA
' Si D sale de zona de sobrecompra VENTA
' El stop (NO CONTROLADO), se pone cuando D vuelve a la zona de la que salio

If Lane_AvisoSZona_Rapido = True Then

   ' Si el valor de Dn > que Dn-1 probamos compra saliendo de zona de sobreventa
   If MatrizTemporal(Mercado, 21, 1) > MatrizTemporal(Mercado, 21, 2) Then
   
      ' Si Dn >= que la zona de Sobreventa y DSn-1 era menor que la zona de SobreVenta
      If MatrizTemporal(Mercado, 21, 1) >= Lane_SobreVenta And MatrizTemporal(Mercado, 21, 2) < Lane_SobreVenta Then
      
         GuardarAvisos Mercado, "COMPRA", "Lane Rapido Estrategia Salida de Zona", "Salida de D (" & CStr(Round(MatrizTemporal(Mercado, 21, 1), 2)) & ") de zona de sobreventa (" & Lane_SobreVenta & ")", "", ""
      
      End If
   
   ' Si el valor de Dn < que Dn-1 probamos compra saliendo de zona de sobrecompra
   ElseIf MatrizTemporal(Mercado, 21, 1) < MatrizTemporal(Mercado, 21, 2) Then
   
      ' Si Dn <= que la zona de Sobrecompra y Dn-1 era mayor que la zona de SobreCompra
      If MatrizTemporal(Mercado, 21, 1) <= Lane_SobreCompra And MatrizTemporal(Mercado, 21, 2) > Lane_SobreCompra Then
      
         GuardarAvisos Mercado, "VENTA", "Lane Rapido Estrategia Salida de Zona", "Salida de D (" & CStr(Round(MatrizTemporal(Mercado, 21, 1), 2)) & ") de zona de sobrecompra (" & Lane_SobreCompra & ")", "", ""
      
      End If
   
   End If

End If

' POPCORN LANE LENTO

' Si DS entra de zona de sobreventa VENTA
' Si DS entra de zona de sobrecompra COMPRA

If Lane_AvisoPopCorn_Lento = True Then

   ' Si el valor de DSn < que DSn-1 probamos venta entrando en zona de sobreventa
   If MatrizTemporal(Mercado, 22, 1) < MatrizTemporal(Mercado, 22, 2) Then
   
      ' Si DSn < que la zona de Sobreventa y DSn-1 era mayor que la zona de SobreVenta
      If MatrizTemporal(Mercado, 22, 1) < Lane_SobreVenta And MatrizTemporal(Mercado, 22, 2) > Lane_SobreVenta Then
      
         GuardarAvisos Mercado, "VENTA", "Lane Lento Estrategia PopCorn", "Entrada de DS (" & CStr(Round(MatrizTemporal(Mercado, 22, 1), 2)) & ") en zona de sobreventa (" & Lane_SobreVenta & ")", "", ""
      
      End If
   
   ' Si el valor de DSn > que DSn-1 probamos compra entrando en zona de sobrecompra
   ElseIf MatrizTemporal(Mercado, 22, 1) > MatrizTemporal(Mercado, 22, 2) Then
   
      ' Si DSn > que la zona de Sobrecompra y DSn-1 era menor que la zona de Sobrecompra
      If MatrizTemporal(Mercado, 22, 1) > Lane_SobreCompra And MatrizTemporal(Mercado, 22, 2) < Lane_SobreCompra Then
      
         GuardarAvisos Mercado, "COMPRA", "Lane Lento Estrategia PopCorn", "Entrada de DS (" & CStr(Round(MatrizTemporal(Mercado, 22, 1), 2)) & ") en zona de sobrecompra (" & Lane_SobreCompra & ")", "", ""
      
      End If
      
   End If
   
End If

' POPCORN LANE RAPIDO

' Si D entra de zona de sobreventa VENTA
' Si D entra de zona de sobrecompra COMPRA

If Lane_AvisoPopCorn_Rapido = True Then

   ' Si el valor de Dn < que Dn-1 probamos venta entrando en zona de sobreventa
   If MatrizTemporal(Mercado, 21, 1) < MatrizTemporal(Mercado, 21, 2) Then
   
      ' Si Dn < que la zona de Sobreventa y Dn-1 era mayor que la zona de SobreVenta
      If MatrizTemporal(Mercado, 21, 1) < Lane_SobreVenta And MatrizTemporal(Mercado, 21, 2) > Lane_SobreVenta Then
      
         GuardarAvisos Mercado, "VENTA", "Lane Rapido Estrategia PopCorn", "Entrada de DS (" & CStr(Round(MatrizTemporal(Mercado, 21, 1), 2)) & ") en zona de sobreventa (" & Lane_SobreVenta & ")", "", ""
      
      End If
   
   ' Si el valor de Dn > que Dn-1 probamos compra entrando en zona de sobrecompra
   ElseIf MatrizTemporal(Mercado, 21, 1) > MatrizTemporal(Mercado, 21, 2) Then
   
      ' Si Dn > que la zona de Sobrecompra y Dn-1 era menor que la zona de Sobrecompra
      If MatrizTemporal(Mercado, 21, 1) > Lane_SobreCompra And MatrizTemporal(Mercado, 21, 2) < Lane_SobreCompra Then
      
         GuardarAvisos Mercado, "COMPRA", "Lane Rapido Estrategia PopCorn", "Entrada de DS (" & CStr(Round(MatrizTemporal(Mercado, 21, 1), 2)) & ") en zona de sobrecompra (" & Lane_SobreCompra & ")", "", ""
      
      End If
      
   End If

End If

End Sub

Public Sub AnalisisTecnicoMercados_Lane(Mercado As Integer)

If MatrizAcciones(((Lane_Periodo - 1) + (Lane_K - 1) + (Lane_D - 1) + (Lane_DS - 1) + (Lane_DSS - 1) + 5), 3) <> "" Then

' 2 + 2 + 2 + 2 + 5 = 13
ReDim MatrizLane((Lane_K - 1) + (Lane_D - 1) + (Lane_DS - 1) + (Lane_DSS - 1) + 5)

' Desde 26 (13 + 2 + 2 + 2 + 2 + 5) hasta 14
For Y = ((Lane_Periodo - 1) + (Lane_K - 1) + (Lane_D - 1) + (Lane_DS - 1) + (Lane_DSS - 1) + 5) To Lane_Periodo Step -1

    MaximoReferencia = 0
    MinimoReferencia = 9999999

    ' Desde 26 hasta 13
    ' ...
    ' Desde 14 hasta 1
    For Z = Y To Y - (Lane_Periodo - 1) Step -1
        
        ' Si el maximo de la cotización es mayor o igual a MaximoReferencia
        If MatrizAcciones(Z, 4) >= MaximoReferencia Then MaximoReferencia = MatrizAcciones(Z, 4)

        ' Si el minimo de la cotización es menor o igual a MinimoReferencia
        If MatrizAcciones(Z, 5) <= MinimoReferencia Then MinimoReferencia = MatrizAcciones(Z, 5)
        
        ' Si estamos en la primera cotización de la referencia
        If Z = Y - (Lane_Periodo - 1) Then
                   
           'MsgBox z & " Min: " & MinimoReferencia & " Max: " & MaximoReferencia & " Cierre: " & MatrizAcciones(z, 3)
                   
           ' Actualizamos el K con el cierre del dia menos el minimo de referencia, dividido entre el maximo de referencia menos el minimo por 100
           
           ' Para evitar problemas de division de 0
           If (CDbl(MatrizAcciones(Z, 3)) - CDbl(MinimoReferencia)) = 0 Then
            
              MatrizLane(Y - (Lane_Periodo - 1)) = 0
              
           Else
           
              MatrizLane(Y - (Lane_Periodo - 1)) = CDbl(((MatrizAcciones(Z, 3) - MinimoReferencia) / (MaximoReferencia - MinimoReferencia)) * 100)
              
           End If
           
        End If
        
    Next
    
Next

' Para aplanamiento K - 1 To 11 (2 + 2 + 2 + 5)
For Y = 1 To (Lane_D - 1) + (Lane_DS - 1) + (Lane_DSS - 1) + 5
    
    ' de 2 a 3
    For Z = Y + 1 To Y + (Lane_K - 1)
    
        MatrizLane(Y) = CDbl(MatrizLane(Y)) + CDbl(MatrizLane(Z))
    
    Next
    
    MatrizLane(Y) = CDbl(MatrizLane(Y)) / CDbl(Lane_K)
  
    If Y <= 5 Then
    
       MatrizTemporal(Mercado, 20, Y) = Round(CDbl(MatrizLane(Y)), 4)
       
    End If

Next

' Para aplanamiento D - 1 To 9 (2 + 2 + 5)
For Y = 1 To (Lane_DS - 1) + (Lane_DSS - 1) + 5
    
    ' de 2 a 3
    For Z = Y + 1 To Y + (Lane_K - 1)
    
        MatrizLane(Y) = CDbl(MatrizLane(Y)) + CDbl(MatrizLane(Z))
    
    Next
    
    MatrizLane(Y) = CDbl(MatrizLane(Y)) / CDbl(Lane_K)
    
    If Y <= 5 Then
    
       MatrizTemporal(Mercado, 21, Y) = Round(CDbl(MatrizLane(Y)), 4)
    
    End If

Next
      
' Para aplanamiento DS - 1 To 7 (2 + 5)
For Y = 1 To (Lane_DSS - 1) + 5
    
    ' de 2 a 3
    For Z = Y + 1 To Y + (Lane_K - 1)
    
        MatrizLane(Y) = CDbl(MatrizLane(Y)) + CDbl(MatrizLane(Z))
    
    Next
    
    MatrizLane(Y) = CDbl(MatrizLane(Y)) / CDbl(Lane_K)
    
    If Y <= 5 Then
    
       MatrizTemporal(Mercado, 22, Y) = Round(CDbl(MatrizLane(Y)), 4)
    
    End If

Next

' Para aplanamiento DSS - 1 To 5
For Y = 1 To 5
    
    ' de 2 a 3
    For Z = Y + 1 To Y + (Lane_K - 1)
    
        MatrizLane(Y) = CDbl(MatrizLane(Y)) + CDbl(MatrizLane(Z))
    
    Next
    
    MatrizLane(Y) = CDbl(MatrizLane(Y)) / CDbl(Lane_K)
    
    If Y <= 5 Then
    
       MatrizTemporal(Mercado, 23, Y) = Round(CDbl(MatrizLane(Y)), 4)
    
    End If

Next
      
      
End If

End Sub
Public Sub CargarDefectos()

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT * FROM Defectos", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   Lane_Periodo = RegistroSQL("Lane_Periodo")
   Lane_K = RegistroSQL("Lane_K")
   Lane_D = RegistroSQL("Lane_D")
   Lane_DS = RegistroSQL("Lane_DS")
   Lane_DSS = RegistroSQL("Lane_DSS")
   Lane_SobreCompra = RegistroSQL("Lane_SobreCompra")
   Lane_SobreVenta = RegistroSQL("Lane_SobreVenta")
   
   RSI_Periodo = RegistroSQL("RSI_Periodo")
   RSI_SobreCompra = RegistroSQL("RSI_SobreCompra")
   RSI_SobreVenta = RegistroSQL("RSI_SobreVenta")

   Lane_AvisoClasica_Lento = RegistroSQL("Lane_AvisoClasica_Lento")

   Lane_AvisoClasica_Rapido = RegistroSQL("Lane_AvisoClasica_Rapido")
   Lane_AvisoSZona_Lento = RegistroSQL("Lane_AvisoSZona_Lento")
   Lane_AvisoSZona_Rapido = RegistroSQL("Lane_AvisoSZona_Rapido")
   Lane_AvisoPopCorn_Lento = RegistroSQL("Lane_AvisoPopCorn_Lento")
   Lane_AvisoPopCorn_Rapido = RegistroSQL("Lane_AvisoPopCorn_Rapido")
   
   RSI_AvisoSalidaZona = RegistroSQL("RSI_AvisoSalidaZona")
   RSI_AvisoFailureSwing = RegistroSQL("RSI_AvisoFailureSwing")
   RSI_AvisoDivergencia = RegistroSQL("RSI_AvisoDivergencia")
   
   ' HAY QUE AGREGARLO A DEFECTOS
   CorreccionOrdenada_Aviso = True
   GapApertura_Aviso = True
   PausaHarami_Aviso = True
   PausaCasiHarami_Aviso = True
   Exceso_Aviso = True
   FuegoPaja_Aviso = True
   Doji_Aviso = True
   Martillo_Aviso = True
   EstrellaFugaz_Aviso = True
   Harami_Aviso = True
   Cobertura_Aviso = True
   Penetrante_Aviso = True
   Envolventes_Aviso = True
   
   Peonza_Aviso = True
   GapTasuki_Aviso = True
   GemelosBlancos_Aviso = True
   LineasSeparacion_Aviso = True
   LineasUnion_Aviso = True
   Puntapie_Aviso = True
   TresRios_Aviso = True

   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close


End Sub



Public Sub CargarTiposVelas()

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT IsNull(MAX(Valor), 0) As Max, IsNull(MIN(Valor), 0) As Min FROM Velas", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim MatrizVelas(CInt(RegistroSQL("Max")), 1)
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Valor, DesCorta FROM Velas", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   MatrizVelas(CInt(RegistroSQL("Valor")), 1) = RegistroSQL("DesCorta")
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub

Public Sub AnalisisTecnicoAcciones()

Dim MaxFecha As Date

MaxFecha = "1/1/1900"

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT COUNT(DISTINCT Acciones.Id_Acciones) As NumeroRegistros FROM Acciones INNER JOIN AccionesMercados ON AccionesMercados.Id_Acciones = Acciones.Id_Acciones INNER JOIN Mercados ON Mercados.Id_Mercados = AccionesMercados.Id_Mercados WHERE Mercados.ControlValorHis = 1", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim MatrizTemporal(RegistroSQL("NumeroRegistros"), 27, 10)
   ReDim MatrizAvisos(RegistroSQL("NumeroRegistros"), 10, 5)
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close
       
VarLinMatrizTemporal = 0
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
   
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Acciones.Id_Acciones, Acciones.Nombre, MAX(ISNULL(AccionesCotizaciones.Fecha, '1/1/1900')) As Fecha, (1) As Generar FROM Acciones INNER JOIN AccionesMercados ON AccionesMercados.Id_Acciones = Acciones.Id_Acciones INNER JOIN Mercados ON Mercados.Id_Mercados = AccionesMercados.Id_Mercados INNER JOIN AccionesCotizaciones ON AccionesCotizaciones.Id_Acciones = Acciones.Id_Acciones LEFT OUTER JOIN AccionesATecnico ON AccionesATecnico.Id_Acciones = Acciones.Id_Acciones WHERE Mercados.ControlValorHis = 1 GROUP BY Mercados.Zona, Acciones.Id_Acciones, Acciones.Nombre ORDER BY Mercados.Zona, Acciones.Nombre, Acciones.Id_Acciones", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   VarLinMatrizTemporal = VarLinMatrizTemporal + 1
   
   MatrizTemporal(VarLinMatrizTemporal, 1, 1) = RegistroSQL("Id_Acciones")
   MatrizTemporal(VarLinMatrizTemporal, 1, 4) = RegistroSQL("Nombre")
   MatrizTemporal(VarLinMatrizTemporal, 1, 2) = RegistroSQL("Fecha")
   MatrizTemporal(VarLinMatrizTemporal, 1, 3) = RegistroSQL("Generar")
   
   ' Ponemos a cero la cantidad de avisos que hay
   MatrizAvisos(VarLinMatrizTemporal, 0, 0) = 0
   
   ' Si la fecha de la ultima cotización es mayor a la varible
   If RegistroSQL("Fecha") > MaxFecha Then MaxFecha = RegistroSQL("Fecha")
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

' Hacemos visible el campo texto del status bar de la pantalla principal
FrmPrincipal.StatusBar1.Panels(3).Visible = True
      
' Recorremos los registros de los mercados
For i = 1 To VarLinMatrizTemporal
 
 ' Si la fecha del valor coincide con la fecha maxima
 If MatrizTemporal(i, 1, 2) = MaxFecha Then
   
   ' Pomemos en el campo texto del status bar el punto en el que vamos
   FrmPrincipal.StatusBar1.Panels(3).Text = CStr(i) & " de " & CStr(VarLinMatrizTemporal) & " - Estudiando Valor " & CStr(MatrizTemporal(i, 1, 4))
    
   DoEvents
   
   ' Desde el segundo grupo hasta el 24
   For A = 2 To 27
   
       ' Desde el primer campo del grupo hasta el 10
       For b = 1 To 10
       
           ' Ponemos todo a 0
           MatrizTemporal(i, A, b) = 0
           
       Next
   
   Next
   
   ' Si el mercado esta marcado como generar
   If MatrizTemporal(i, 1, 3) = 1 Then

      ' 1310 porque son el número maximo de cotizaciones si tenemos 5 años completos
      ReDim MatrizAcciones(1310, 10)
      
      'Erase MatrizAcciones
      VarLinMatrizAcciones = 0
   
      ' Abrimos la base de datos
      ConexionSQL.Open FicheroUDLSQL

      ' Abrimos el recordset de la consulta
      RegistroSQL.Open "SELECT TOP 1310 Fecha, Apertura, Cierre, Maximo, Minimo, Volumen FROM AccionesCotizaciones WHERE Id_Acciones = '" & MatrizTemporal(i, 1, 1) & "' AND YEAR(Fecha) >= '" & CStr(Year(MatrizTemporal(i, 1, 2)) - 4) & "'ORDER BY Fecha DESC", ConexionSQL, adOpenDynamic, adLockOptimistic

      Do While Not RegistroSQL.EOF
      
         VarLinMatrizAcciones = VarLinMatrizAcciones + 1
      
         MatrizAcciones(VarLinMatrizAcciones, 1) = RegistroSQL("Fecha")
         MatrizAcciones(VarLinMatrizAcciones, 2) = RegistroSQL("Apertura")
         MatrizAcciones(VarLinMatrizAcciones, 3) = RegistroSQL("Cierre")
         MatrizAcciones(VarLinMatrizAcciones, 4) = RegistroSQL("Maximo")
         MatrizAcciones(VarLinMatrizAcciones, 5) = RegistroSQL("Minimo")
         MatrizAcciones(VarLinMatrizAcciones, 6) = RegistroSQL("Volumen")
      
         RegistroSQL.MoveNext
    
      Loop

      ' Cerramos los objetos abiertos para la conexión
      RegistroSQL.Close
      ConexionSQL.Close
      
      AnalisisTecnicoMercados_MediasMoviles (i)
      
      AnalisisTecnicoMercados_Resto (i)
      
      AnalisisTecnicoMercados_TipoVela (i)
      
      AnalisisTecnicoMercados_Lane (i)
      
      AnalisisTecnicoMercados_RSI_Suavizado (i)
      
      AnalisisEstrategias (i)
      
      GuardarAvisosAcciones (i)
      
      AnalisisTecnicoAcciones_ActTabla_Indicadores (i)
      
      AnalisisTecnicoAcciones_ActTabla (i)
      
      Erase MatrizAcciones
   
   End If
   
 End If
   
Next

Erase MatrizTemporal

' Hacemos invisible el campo texto del status bar de la pantalla principal
FrmPrincipal.StatusBar1.Panels(3).Visible = False

End Sub

' ANÁLISIS TÉCNICO DE LOS MERCADOS
'
' De los mercados que tengan marcado el control histórico

Public Sub AnalisisTecnicoMercados()

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT COUNT(Mercados.Id_Mercados) As NumeroRegistros FROM Mercados WHERE Mercados.ControlHis = 1", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim MatrizTemporal(RegistroSQL("NumeroRegistros"), 27, 10)
   ReDim MatrizAvisos(RegistroSQL("NumeroRegistros"), 10, 5)
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close
       
VarLinMatrizTemporal = 0
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
   
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Mercados.Id_Mercados, Mercados.Nombre, MAX(ISNULL(MercadosCotizaciones.Fecha, '1/1/1900')) As Fecha, (CASE WHEN MAX(ISNULL(MercadosCotizaciones.Fecha, '1/1/1900')) = MAX(ISNULL(MercadosATecnico.FechaDatos, '1/1/1900')) THEN 0 ELSE 1 END) As Generar FROM Mercados INNER JOIN MercadosCotizaciones ON MercadosCotizaciones.Id_Mercados = Mercados.Id_Mercados LEFT OUTER JOIN MercadosATecnico ON MercadosATecnico.Id_Mercados = Mercados.Id_Mercados WHERE Mercados.ControlHis = 1 GROUP BY Mercados.Id_Mercados, Mercados.Nombre ORDER BY Mercados.Id_Mercados, Mercados.Nombre", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   VarLinMatrizTemporal = VarLinMatrizTemporal + 1
   
   MatrizTemporal(VarLinMatrizTemporal, 1, 1) = RegistroSQL("Id_Mercados")
   MatrizTemporal(VarLinMatrizTemporal, 1, 4) = RegistroSQL("Nombre")
   MatrizTemporal(VarLinMatrizTemporal, 1, 2) = RegistroSQL("Fecha")
   MatrizTemporal(VarLinMatrizTemporal, 1, 3) = RegistroSQL("Generar")
   
   ' Ponemos a cero la cantidad de avisos que hay
   MatrizAvisos(VarLinMatrizTemporal, 0, 0) = 0

   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

' Hacemos visible el campo texto del status bar de la pantalla principal
FrmPrincipal.StatusBar1.Panels(3).Visible = True
      

' Recorremos los registros de los mercados
For i = 1 To VarLinMatrizTemporal
 
   ' Pomemos en el campo texto del status bar el punto en el que vamos
   FrmPrincipal.StatusBar1.Panels(3).Text = CStr(i) & " de " & CStr(VarLinMatrizTemporal) & " - Estudiando Mercado " & CStr(MatrizTemporal(i, 1, 4))
    
   ' Desde el segundo grupo hasta el 24
   For A = 2 To 27
   
       ' Desde el primer campo del grupo hasta el 10
       For b = 1 To 10
       
           ' Ponemos todo a 0
           MatrizTemporal(i, A, b) = 0
           
       Next
   
   Next
   
   ' Si el mercado esta marcado como generar
   If MatrizTemporal(i, 1, 3) = 1 Then

      ' 1310 porque son el número maximo de cotizaciones si tenemos 5 años completos
      ReDim MatrizAcciones(1310, 10)
      
      'Erase MatrizAcciones
      VarLinMatrizAcciones = 0
   
      ' Abrimos la base de datos
      ConexionSQL.Open FicheroUDLSQL

      ' Abrimos el recordset de la consulta
      RegistroSQL.Open "SELECT TOP 1310 Fecha, Apertura, Cierre, Maximo, Minimo, Volumen FROM MercadosCotizaciones WHERE Id_Mercados = '" & MatrizTemporal(i, 1, 1) & "' AND YEAR(Fecha) >= '" & CStr(Year(MatrizTemporal(i, 1, 2)) - 4) & "'ORDER BY Fecha DESC", ConexionSQL, adOpenDynamic, adLockOptimistic

      Do While Not RegistroSQL.EOF
      
         VarLinMatrizAcciones = VarLinMatrizAcciones + 1
      
         MatrizAcciones(VarLinMatrizAcciones, 1) = RegistroSQL("Fecha")
         MatrizAcciones(VarLinMatrizAcciones, 2) = RegistroSQL("Apertura")
         MatrizAcciones(VarLinMatrizAcciones, 3) = RegistroSQL("Cierre")
         MatrizAcciones(VarLinMatrizAcciones, 4) = RegistroSQL("Maximo")
         MatrizAcciones(VarLinMatrizAcciones, 5) = RegistroSQL("Minimo")
         MatrizAcciones(VarLinMatrizAcciones, 6) = RegistroSQL("Volumen")
      
         RegistroSQL.MoveNext
    
      Loop

      ' Cerramos los objetos abiertos para la conexión
      RegistroSQL.Close
      ConexionSQL.Close
      
      AnalisisTecnicoMercados_MediasMoviles (i)
      
      AnalisisTecnicoMercados_Resto (i)
      
      AnalisisTecnicoMercados_TipoVela (i)
      
      AnalisisTecnicoMercados_Lane (i)
      
      AnalisisTecnicoMercados_RSI (i)
      
      AnalisisEstrategias (i)
      
      GuardarAvisosMercados (i)
      
      AnalisisTecnicoMercados_ActTabla_Indicadores (i)
      
      AnalisisTecnicoMercados_ActTabla (i)
      
      Erase MatrizAcciones
   
   End If
  
Next

Erase MatrizTemporal

' Hacemos invisible el campo texto del status bar de la pantalla principal
FrmPrincipal.StatusBar1.Panels(3).Visible = False

End Sub

Public Sub AnalisisTecnicoMercados_TipoVela(Mercado As Integer)

CuerpoCotizacion = MatrizTemporal(Mercado, 10, 3) - MatrizTemporal(Mercado, 10, 2)

'MsgBox CuerpoCotizacion

' Recorremos las ultimas 200 cotizaciones
For Y = 1 To 200

    ' Si la apertura, cierre, maximo y minimo es igual
    If MatrizAcciones(Y, 2) = MatrizAcciones(Y, 3) And MatrizAcciones(Y, 4) = MatrizAcciones(Y, 5) And MatrizAcciones(Y, 2) = MatrizAcciones(Y, 4) Then
       
       ' Marcamos la vela inexistente
       MatrizAcciones(Y, 7) = 0
    
    ' Si la apertura es igual al cierre es un doji puro
    ElseIf MatrizAcciones(Y, 2) = MatrizAcciones(Y, 3) Then
    
       ' El recorrido de la cotización del dia es el maximo menos el minimo
       RecorridoCotizacion = MatrizAcciones(Y, 4) - MatrizAcciones(Y, 5)
       
       ' El recorrido lo dividimos entre tres para ver donde esta la el cierre
       RecorridoCotizacion = RecorridoCotizacion * 0.3333
    
       ' Si el cierre es igual al maximo es un doji dragon volador
       If MatrizAcciones(Y, 3) = MatrizAcciones(Y, 4) Then
          
          ' Marcamos la vela como doji dragon volador
          MatrizAcciones(Y, 7) = 4
          
       ' Si el cierre es igual al minimo es un doji piedra funeraria
       ElseIf MatrizAcciones(Y, 3) = MatrizAcciones(Y, 5) Then
          
          ' Marcamos la vela como doji piedra funeraria
          MatrizAcciones(Y, 7) = 5
          
       ' Si el minimo, maximo, apertura y cierre son lo mismo
       ElseIf RecorridoCotizacion = 0 Then
          
          ' Marcamos la vela como doji mas
          MatrizAcciones(Y, 7) = 1
          
       ' Si el cierre es menor o igual al minimo más un tercio del recorrido
       ElseIf MatrizAcciones(Y, 3) <= (MatrizAcciones(Y, 5) + RecorridoCotizacion) Then
          
          ' Marcamos la vela como doji cruz invertida
          MatrizAcciones(Y, 7) = 3
       
       ' Si el cierre es mayor o igual al maximo menos un tercio del recorrido
       ElseIf MatrizAcciones(Y, 3) >= (MatrizAcciones(Y, 4) - RecorridoCotizacion) Then
          
          ' Marcamos la vela como doji cruz
          MatrizAcciones(Y, 7) = 2
          
       ' Si no encaja en ninguno de los anteriores supestos es un doji mas
       Else
       
          ' Marcamos la vela como doji mas
           MatrizAcciones(Y, 7) = 1
          
       End If
       
    'Si la diferencia entre la apertura y el cierre no supera un 0,15% del valor de cierre de la cotizacion
    ElseIf (Abs(MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3)) / (MatrizAcciones(Y, 3) * 0.01)) < 0.15 Then
    
       ' El recorrido de la cotización del dia es el maximo menos el minimo
       RecorridoCotizacion = MatrizAcciones(Y, 4) - MatrizAcciones(Y, 5)
       
       ' El recorrido lo dividimos entre tres para ver donde esta la el cierre
       RecorridoCotizacion = RecorridoCotizacion * 0.3333
    
       ' Si el cierre es igual al maximo es un doji dragon volador
       If MatrizAcciones(Y, 3) = MatrizAcciones(Y, 4) Then
          
          ' Marcamos la vela como doji dragon volador
          MatrizAcciones(Y, 7) = 4
          
       ' Si el cierre es igual al minimo es un doji piedra funeraria
       ElseIf MatrizAcciones(Y, 3) = MatrizAcciones(Y, 5) Then
          
          ' Marcamos la vela como doji piedra funeraria
          MatrizAcciones(Y, 7) = 5
          
       ' Si el minimo, maximo, apertura y cierre son lo mismo
       ElseIf RecorridoCotizacion = 0 Then
          
          ' Marcamos la vela como doji mas
          MatrizAcciones(Y, 7) = 1
          
       ' Si el cierre es menor o igual al minimo más un tercio del recorrido
       ElseIf MatrizAcciones(Y, 3) <= (MatrizAcciones(Y, 5) + RecorridoCotizacion) Then
          
          ' Marcamos la vela como doji cruz invertida
          MatrizAcciones(Y, 7) = 3
       
       ' Si el cierre es mayor o igual al maximo menos un tercio del recorrido
       ElseIf MatrizAcciones(Y, 3) >= (MatrizAcciones(Y, 4) - RecorridoCotizacion) Then
          
          ' Marcamos la vela como doji cruz
          MatrizAcciones(Y, 7) = 2
          
       ' Si no encaja en ninguno de los anteriores supestos es un doji mas
       Else
       
          ' Marcamos la vela como doji mas
           MatrizAcciones(Y, 7) = 1
          
       End If
       
       'Si la apertura es mayor que el cierre
       If MatrizAcciones(Y, 2) > MatrizAcciones(Y, 3) Then
        
          ' La vela es negra
          MatrizAcciones(Y, 7) = MatrizAcciones(Y, 7) * -1
       
       End If
    
    ' Si la diferencia entre la apertura y el cierre es menor o igual que la vela media 200 - 10% del cuerpo entero vela
    ' CUERPO PEQUEÑO
    ElseIf (Abs(MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3))) <= (MatrizTemporal(Mercado, 10, 1) - (CuerpoCotizacion * 0.1)) Then
       
       ' El recorrido de la cotización del dia es el maximo menos el minimo
       RecorridoCotizacion = MatrizAcciones(Y, 4) - MatrizAcciones(Y, 5)
       
       ' El recorrido lo dividimos entre tres para ver donde esta la el cierre
       RecorridoCotizacion = RecorridoCotizacion * 0.4
       
       ' Si la apertura y el cierre son mayores al maximo menos el 40% del recorrido del dia
       If (MatrizAcciones(Y, 2) > (MatrizAcciones(Y, 4) - RecorridoCotizacion)) And (MatrizAcciones(Y, 3) > (MatrizAcciones(Y, 4) - RecorridoCotizacion)) Then
          
          ' Marcamos la vela como martillo/colgado
          MatrizAcciones(Y, 7) = 11
          
       ' Si la apertura y el cierre son menores al minimo mas el 40% del recorrido del dia
       ElseIf (MatrizAcciones(Y, 2) < (MatrizAcciones(Y, 5) + RecorridoCotizacion)) And (MatrizAcciones(Y, 3) < (MatrizAcciones(Y, 5) + RecorridoCotizacion)) Then
          
          ' Marcamos la vela como estrella fugaz
          MatrizAcciones(Y, 7) = 12
          
       Else
       
          ' Marcamos la vela como peonza
          MatrizAcciones(Y, 7) = 13
       
       End If
       
       ' Si la apertura es mayor que el cierre
       If MatrizAcciones(Y, 2) > MatrizAcciones(Y, 3) Then
          
          ' Marcamos la vela como negra
          MatrizAcciones(Y, 7) = MatrizAcciones(Y, 7) * -1
          
       
       End If
       
   ' Si la diferencia entre la apertura y el cierre es mayor o igual que la vela media 200 + 10% del cuerpo entero vela
    ' CUERPO GRANDE
    ElseIf (Abs(MatrizAcciones(Y, 2) - MatrizAcciones(Y, 3))) >= (MatrizTemporal(Mercado, 10, 1) + (CuerpoCotizacion * 0.1)) Then
    
       ' Si la apertura es igual al minimo y el cierre igual al maximo
       If MatrizAcciones(Y, 2) = MatrizAcciones(Y, 5) And MatrizAcciones(Y, 3) = MatrizAcciones(Y, 4) Then
       
          ' Marcamos la vela como marubozu grande blanco
          MatrizAcciones(Y, 7) = 31
       
       ' Si la apertura es igual al maximo y el cierre igual al minimo
       ElseIf MatrizAcciones(Y, 2) = MatrizAcciones(Y, 4) And MatrizAcciones(Y, 3) = MatrizAcciones(Y, 5) Then
       
          ' Marcamos la vela como marubozu grande negro
          MatrizAcciones(Y, 7) = -31
       
       Else
       
          ' Si la apertura es mayor al cierre
          If MatrizAcciones(Y, 2) > MatrizAcciones(Y, 3) Then
          
             ' Marcamos la vela como cuerpo grande negro
             MatrizAcciones(Y, 7) = -32
          
          Else
          
             ' Marcamos la vela como cuerpo grande blanco
             MatrizAcciones(Y, 7) = 32
          
          End If
       
          
          
       End If
       
    ' Solo queda CUERPO MEDIANO
    Else
    
       ' Si la apertura es igual al minimo y el cierre igual al maximo
       If MatrizAcciones(Y, 2) = MatrizAcciones(Y, 5) And MatrizAcciones(Y, 3) = MatrizAcciones(Y, 4) Then
       
          ' Marcamos la vela como marubozu mediano blanco
          MatrizAcciones(Y, 7) = 21
       
       ' Si la apertura es igual al maximo y el cierre igual al minimo
       ElseIf MatrizAcciones(Y, 2) = MatrizAcciones(Y, 4) And MatrizAcciones(Y, 3) = MatrizAcciones(Y, 5) Then
       
          ' Marcamos la vela como marubozu mediano negro
          MatrizAcciones(Y, 7) = -21
       
       Else
       
          ' Si la apertura es mayor al cierre
          If MatrizAcciones(Y, 2) > MatrizAcciones(Y, 3) Then
          
             ' Marcamos la vela como cuerpo mediano negro
             MatrizAcciones(Y, 7) = -22
          
          Else
          
             ' Marcamos la vela como cuerpo mediano blanco
             MatrizAcciones(Y, 7) = 22
          
          End If
       
          
          
       End If
    
    End If

Next

         
End Sub

Public Sub AnalisisTecnicoMercados_Resto(Mercado As Integer)

     ' PARA EL CONTROL DE RESISTENCIA (MAXIMO 5 AÑOS) Y SOPORTE (MINIMO 5 AÑOS)
      
         
      'Para el control de soportes y resistencias
      FSoporte = ""
      CSoporte = 9999999
      FResistencia = ""
      CResistencia = 0
         
      'Para controlar la posicion en la matriz
      PosSoporte = 0
      PosResistencia = 0
      
      'Para controlar la posicion en la matriz
      PosSoporteUltimo = 0
      PosResistenciaUltima = 0
      
      ' Recorremos los volumenes
      For Y = 1310 To 1 Step -1
      
          ' Si el registro tiene datos de cotización
          If MatrizAcciones(Y, 3) <> "" Then

             ' Si la cotización del cierre es menor o igual a la cotización del soporte del perido
             If CDbl(MatrizAcciones(Y, 3)) <= CDbl(CSoporte) Then
             
                FSoporte = MatrizAcciones(Y, 1)
                CSoporte = MatrizAcciones(Y, 3)
                PosSoporte = Y
             
             End If
             
             ' Si la cotización del cierre es mayor o igual a la cotización de la resitencia del perido
             If CDbl(MatrizAcciones(Y, 3)) >= CDbl(CResistencia) Then
             
                FResistencia = MatrizAcciones(Y, 1)
                CResistencia = MatrizAcciones(Y, 3)
                PosResistencia = Y
             
             End If
          
          End If
             
      Next
         
      'Si se ha marcado algun soporte
      If PosSoporte <> 0 Then
            
         DifSoporte = 9999999
         
         For Z = PosSoporte - 1 To 1 Step -1
            
             ' Si la cotizacion menos la cotizacion del soporte dividido entre los dias es menor a la diferencia anterior
             If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z)) <= DifSoporte Then
                
                DifSoporte = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z))
                
                PosSoporteUltimo = Z
                
             End If
            
         Next
            
         ' Actualizamos datos tendencia alcista 5 años
         MatrizTemporal(Mercado, 14, 1) = FSoporte
         MatrizTemporal(Mercado, 14, 2) = CSoporte
            
         If MatrizAcciones(PosSoporteUltimo, 1) = "" Then
                
            MatrizTemporal(Mercado, 14, 3) = FSoporte
            
         Else
               
            MatrizTemporal(Mercado, 14, 3) = MatrizAcciones(PosSoporteUltimo, 1)
               
         End If
            
         If MatrizAcciones(PosSoporteUltimo, 3) = "" Then
                
            MatrizTemporal(Mercado, 14, 4) = CSoporte
            
         Else
               
            MatrizTemporal(Mercado, 14, 4) = MatrizAcciones(PosSoporteUltimo, 3)
               
         End If
            
         If DifSoporte = 9999999 Then
                
            MatrizTemporal(Mercado, 14, 5) = 0
            
         Else
               
            MatrizTemporal(Mercado, 14, 5) = Round(CDbl(DifSoporte) / (CDbl(CSoporte) * 0.01), 4)
               
         End If
            
         MatrizTemporal(Mercado, 14, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CSoporte)) / (CDbl(CSoporte) * 0.01), 4)
         MatrizTemporal(Mercado, 14, 7) = PosSoporte - 1
        
      End If
         
      'Si se ha marcado alguna resistencia
      If PosResistencia <> 0 Then
            
         DifResistencia = -9999999
         
         For Z = PosResistencia - 1 To 1 Step -1
            
             ' Si la cotizacion menos la cotizacion de la resistencia dividido entre los dias es mayor a la diferencia anterior
             If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z)) >= DifResistencia Then
                
                DifResistencia = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z))
                
                PosResistenciaUltima = Z
                   
             End If
            
         Next
            
         ' Actualizamos datos tendencia bajista 200 dias
         MatrizTemporal(Mercado, 15, 1) = FResistencia
         MatrizTemporal(Mercado, 15, 2) = CResistencia
            
         If MatrizAcciones(PosResistenciaUltima, 1) = "" Then
               
            MatrizTemporal(Mercado, 15, 3) = FResistencia
            
         Else
               
            MatrizTemporal(Mercado, 15, 3) = MatrizAcciones(PosResistenciaUltima, 1)
               
         End If
            
         If MatrizAcciones(PosResistenciaUltima, 3) = "" Then
                
            MatrizTemporal(Mercado, 15, 4) = CResistencia
            
         Else
               
            MatrizTemporal(Mercado, 15, 4) = MatrizAcciones(PosResistenciaUltima, 3)
               
         End If
            
         If DifResistencia = -9999999 Then
                
            MatrizTemporal(Mercado, 15, 5) = 0
            
         Else
               
            MatrizTemporal(Mercado, 15, 5) = Round(CDbl(DifResistencia) / (CDbl(CResistencia) * 0.01), 4)
              
         End If
            
         MatrizTemporal(Mercado, 15, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CResistencia)) / (CDbl(CResistencia) * 0.01), 4)
         MatrizTemporal(Mercado, 15, 7) = PosResistencia - 1
        
      End If
      
      
      ' PARA EL CONTROL DE MIN, MAX, MEDIA VOLUMEN 200 DIAS Y MIN, MAX, MEDIA VELA 200 DIAS
      ' RESISTENCIA (MAXIMO 200 DIAS) Y SOPORTE (MINIMO 200 DIAS)
      
      ' Si el ultimo registro es númerico entendemos que el resto tambien
      If MatrizAcciones(200, 6) <> "" Then
         
         'Para el control de soportes y resistencias
         FSoporte = ""
         CSoporte = 9999999
         FResistencia = ""
         CResistencia = 0
         
         'Para controlar la posicion en la matriz
         PosSoporte = 0
         PosResistencia = 0
      
         'Para controlar la posicion en la matriz
         PosSoporteUltimo = 0
         PosResistenciaUltima = 0
      
         ' Recorremos los volumenes
         For Y = 200 To 1 Step -1
         
             ' Si el volumen minimo del mercado es 0 o es mayor que el volumen del día
             If MatrizAcciones(Y, 6) < MatrizTemporal(Mercado, 7, 2) Or MatrizTemporal(Mercado, 7, 2) = 0 Then
             
                MatrizTemporal(Mercado, 7, 2) = MatrizAcciones(Y, 6)
             
             End If
             
             ' Si el volumen maximo del mercado es menor que el volumen del día
             If MatrizAcciones(Y, 6) > MatrizTemporal(Mercado, 7, 3) Then
             
                MatrizTemporal(Mercado, 7, 3) = MatrizAcciones(Y, 6)
             
             End If
              
             ' Agregamos a la medía de volumen el volumen del día
             MatrizTemporal(Mercado, 7, 1) = MatrizTemporal(Mercado, 7, 1) + MatrizAcciones(Y, 6)
              
             ' Si la vela minima del mercado es 0 o es mayor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) < MatrizTemporal(Mercado, 10, 2) Or MatrizTemporal(Mercado, 10, 2) = 0 Then
             
                MatrizTemporal(Mercado, 10, 2) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
             
             ' Si la vela maxima del mercado es menor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) > MatrizTemporal(Mercado, 10, 3) Then
             
                MatrizTemporal(Mercado, 10, 3) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
              
             ' Agregamos a la medía de vela la vela del día
             MatrizTemporal(Mercado, 10, 1) = MatrizTemporal(Mercado, 10, 1) + Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
              
             ' Si la cotización del cierre es menor o igual a la cotización del soporte del perido
             If CDbl(MatrizAcciones(Y, 3)) <= CDbl(CSoporte) Then
             
                FSoporte = MatrizAcciones(Y, 1)
                CSoporte = MatrizAcciones(Y, 3)
                PosSoporte = Y
             
             End If
             
             ' Si la cotización del cierre es mayor o igual a la cotización de la resitencia del perido
             If CDbl(MatrizAcciones(Y, 3)) >= CDbl(CResistencia) Then
             
                FResistencia = MatrizAcciones(Y, 1)
                CResistencia = MatrizAcciones(Y, 3)
                PosResistencia = Y
             
             End If
             
         Next
          
         ' Ajustamos el total de volumen entre total de dias
         MatrizTemporal(Mercado, 7, 1) = MatrizTemporal(Mercado, 7, 1) / 200
               
          
         ' Ajustamos el total de velas entre el total de días
         MatrizTemporal(Mercado, 10, 1) = MatrizTemporal(Mercado, 10, 1) / 200
         
         ' Actualizamos cotizaciones y fechas de soporte y resistencia
         MatrizTemporal(Mercado, 13, 1) = CSoporte
         MatrizTemporal(Mercado, 13, 2) = FSoporte
         MatrizTemporal(Mercado, 13, 3) = CResistencia
         MatrizTemporal(Mercado, 13, 4) = FResistencia
         
         'Si se ha marcado algun soporte
         If PosSoporte <> 0 Then
            
            DifSoporte = 9999999
         
            For Z = PosSoporte - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion del soporte dividido entre los dias es menor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z)) <= DifSoporte Then
                
                   DifSoporte = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z))
                
                   PosSoporteUltimo = Z
                
                End If
            
            Next
            
            ' Actualizamos datos tendencia alcista 200 dias
            MatrizTemporal(Mercado, 16, 1) = FSoporte
            MatrizTemporal(Mercado, 16, 2) = CSoporte
            
            If MatrizAcciones(PosSoporteUltimo, 1) = "" Then
                
               MatrizTemporal(Mercado, 16, 3) = FSoporte
            
            Else
               
               MatrizTemporal(Mercado, 16, 3) = MatrizAcciones(PosSoporteUltimo, 1)
               
            End If
            
            If MatrizAcciones(PosSoporteUltimo, 3) = "" Then
                
               MatrizTemporal(Mercado, 16, 4) = CSoporte
            
            Else
               
               MatrizTemporal(Mercado, 16, 4) = MatrizAcciones(PosSoporteUltimo, 3)
               
            End If
            
            If DifSoporte = 9999999 Then
                
               MatrizTemporal(Mercado, 16, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 16, 5) = Round(CDbl(DifSoporte) / (CDbl(CSoporte) * 0.01), 4)
               
            End If
            
            MatrizTemporal(Mercado, 16, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CSoporte)) / (CDbl(CSoporte) * 0.01), 4)
            MatrizTemporal(Mercado, 16, 7) = PosSoporte - 1
        
         End If
         
         'Si se ha marcado alguna resistencia
         If PosResistencia <> 0 Then
            
            DifResistencia = -9999999
         
            For Z = PosResistencia - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion de la resistencia dividido entre los dias es mayor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z)) >= DifResistencia Then
                
                   DifResistencia = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z))
                
                   PosResistenciaUltima = Z
                   
                End If
            
            Next
            
            ' Actualizamos datos tendencia bajista 200 dias
            MatrizTemporal(Mercado, 17, 1) = FResistencia
            MatrizTemporal(Mercado, 17, 2) = CResistencia
            
            If MatrizAcciones(PosResistenciaUltima, 1) = "" Then
                
               MatrizTemporal(Mercado, 17, 3) = FResistencia
            
            Else
               
               MatrizTemporal(Mercado, 17, 3) = MatrizAcciones(PosResistenciaUltima, 1)
               
            End If
            
            If MatrizAcciones(PosResistenciaUltima, 3) = "" Then
                
               MatrizTemporal(Mercado, 17, 4) = CResistencia
            
            Else
               
               MatrizTemporal(Mercado, 17, 4) = MatrizAcciones(PosResistenciaUltima, 3)
               
            End If
            
            If DifResistencia = -9999999 Then
                
               MatrizTemporal(Mercado, 17, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 17, 5) = Round(CDbl(DifResistencia) / (CDbl(CResistencia) * 0.01), 4)
               
            End If
            
            MatrizTemporal(Mercado, 17, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CResistencia)) / (CDbl(CResistencia) * 0.01), 4)
            MatrizTemporal(Mercado, 17, 7) = PosResistencia - 1
        
         End If
         
         
      End If
      
      ' PARA EL CONTROL DE MIN, MAX, MEDIA VOLUMEN 50 DIAS Y MIN, MAX, MEDIA VELA 50 DIAS
      ' RESISTENCIA (MAXIMO 50 DIAS) Y SOPORTE (MINIMO 50 DIAS)
      
      ' Si el ultimo registro es númerico entendemos que el resto tambien
      If MatrizAcciones(50, 6) <> "" Then
      
         'Para el control de soportes y resistencias
         FSoporte = ""
         CSoporte = 9999999
         FResistencia = ""
         CResistencia = 0
         
         ' Para controlar la posicion en la matriz
         PosSoporte = 0
         PosResistencia = 0
      
         'Para controlar la posicion en la matriz
         PosSoporteUltimo = 0
         PosResistenciaUltima = 0
      
         ' Recorremos los volumenes
         For Y = 50 To 1 Step -1
         
             ' Si el volumen minimo del mercado es 0 o es mayor que el volumen del día
             If MatrizAcciones(Y, 6) < MatrizTemporal(Mercado, 6, 2) Or MatrizTemporal(Mercado, 6, 2) = 0 Then
             
                MatrizTemporal(Mercado, 6, 2) = MatrizAcciones(Y, 6)
             
             End If
             
             ' Si el volumen maximo del mercado es menor que el volumen del día
             If MatrizAcciones(Y, 6) > MatrizTemporal(Mercado, 6, 3) Then
             
                MatrizTemporal(Mercado, 6, 3) = MatrizAcciones(Y, 6)
             
             End If
              
             ' Agregamos a la medía de volumen el volumen del día
             MatrizTemporal(Mercado, 6, 1) = MatrizTemporal(Mercado, 6, 1) + MatrizAcciones(Y, 6)
              
             ' Si la vela minima del mercado es 0 o es mayor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) < MatrizTemporal(Mercado, 9, 2) Or MatrizTemporal(Mercado, 9, 2) = 0 Then
             
                MatrizTemporal(Mercado, 9, 2) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
             
             ' Si la vela maxima del mercado es menor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) > MatrizTemporal(Mercado, 9, 3) Then
             
                MatrizTemporal(Mercado, 9, 3) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
              
             ' Agregamos a la medía de vela la vela del día
             MatrizTemporal(Mercado, 9, 1) = MatrizTemporal(Mercado, 9, 1) + Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
              
             ' Si la cotización del cierre es menor o igual a la cotización del soporte del perido
             If CDbl(MatrizAcciones(Y, 3)) <= CDbl(CSoporte) Then
             
                FSoporte = MatrizAcciones(Y, 1)
                CSoporte = MatrizAcciones(Y, 3)
                PosSoporte = Y
             
             End If
             
             ' Si la cotización del cierre es mayor o igual a la cotización de la resitencia del perido
             If CDbl(MatrizAcciones(Y, 3)) >= CDbl(CResistencia) Then
             
                FResistencia = MatrizAcciones(Y, 1)
                CResistencia = MatrizAcciones(Y, 3)
                PosResistencia = Y
             
             End If
             
         Next
         
         ' Ajustamos el total de volumen entre total de dias
         MatrizTemporal(Mercado, 6, 1) = MatrizTemporal(Mercado, 6, 1) / 50
          
         ' Ajustamos el total de velas entre el total de días
         MatrizTemporal(Mercado, 9, 1) = MatrizTemporal(Mercado, 9, 1) / 50
          
         ' Actualizamos cotizaciones y fechas de soporte y resistencia
         MatrizTemporal(Mercado, 12, 1) = CSoporte
         MatrizTemporal(Mercado, 12, 2) = FSoporte
         MatrizTemporal(Mercado, 12, 3) = CResistencia
         MatrizTemporal(Mercado, 12, 4) = FResistencia
         
                  'Si se ha marcado algun soporte
         If PosSoporte <> 0 Then
            
            DifSoporte = 9999999
         
            For Z = PosSoporte - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion del soporte dividido entre los dias es menor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z)) <= DifSoporte Then
                
                   DifSoporte = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z))
                
                   PosSoporteUltimo = Z
                
                End If
            
            Next
            
            ' Actualizamos datos tendencia alcista 200 dias
            MatrizTemporal(Mercado, 18, 1) = FSoporte
            MatrizTemporal(Mercado, 18, 2) = CSoporte
            
            If MatrizAcciones(PosSoporteUltimo, 1) = "" Then
                
               MatrizTemporal(Mercado, 18, 3) = FSoporte
            
            Else
               
               MatrizTemporal(Mercado, 18, 3) = MatrizAcciones(PosSoporteUltimo, 1)
               
            End If
            
            If MatrizAcciones(PosSoporteUltimo, 3) = "" Then
                
               MatrizTemporal(Mercado, 18, 4) = CSoporte
            
            Else
               
               MatrizTemporal(Mercado, 18, 4) = MatrizAcciones(PosSoporteUltimo, 3)
               
            End If
            
            If DifSoporte = 9999999 Then
                
               MatrizTemporal(Mercado, 18, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 18, 5) = Round(CDbl(DifSoporte) / (CDbl(CSoporte) * 0.01), 4)
               
            End If
            
            MatrizTemporal(Mercado, 18, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CSoporte)) / (CDbl(CSoporte) * 0.01), 4)
            MatrizTemporal(Mercado, 18, 7) = PosSoporte - 1
        
         End If
         
         'Si se ha marcado alguna resistencia
         If PosResistencia <> 0 Then
            
            DifResistencia = -9999999
         
            For Z = PosResistencia - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion de la resistencia dividido entre los dias es mayor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z)) >= DifResistencia Then
                
                   DifResistencia = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z))
                
                   PosResistenciaUltima = Z
                   
                End If
            
            Next
            
            ' Actualizamos datos tendencia bajista 200 dias
            MatrizTemporal(Mercado, 19, 1) = FResistencia
            MatrizTemporal(Mercado, 19, 2) = CResistencia
            
            If MatrizAcciones(PosResistenciaUltima, 1) = "" Then
                
               MatrizTemporal(Mercado, 19, 3) = FResistencia
            
            Else
               
               MatrizTemporal(Mercado, 19, 3) = MatrizAcciones(PosResistenciaUltima, 1)
               
            End If
            
            If MatrizAcciones(PosResistenciaUltima, 3) = "" Then
                
               MatrizTemporal(Mercado, 19, 4) = CResistencia
            
            Else
               
               MatrizTemporal(Mercado, 19, 4) = MatrizAcciones(PosResistenciaUltima, 3)
               
            End If
            
            If DifResistencia = -9999999 Then
                
               MatrizTemporal(Mercado, 19, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 19, 5) = Round(CDbl(DifResistencia) / (CDbl(CResistencia) * 0.01), 4)
               
            End If
            
            MatrizTemporal(Mercado, 19, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CResistencia)) / (CDbl(CResistencia) * 0.01), 4)
            MatrizTemporal(Mercado, 19, 7) = PosResistencia - 1
        
         End If
         
      End If
      
      ' PARA EL CONTROL DE MIN, MAX, MEDIA VOLUMEN 20 DIAS Y MIN, MAX, MEDIA VELA 20 DIAS
      ' RESISTENCIA (MAXIMO 20 DIAS) Y SOPORTE (MINIMO 20 DIAS)
      
      ' Si el ultimo registro es númerico entendemos que el resto tambien
      If MatrizAcciones(20, 6) <> "" Then
      
         'Para el control de soportes y resistencias
         FSoporte = ""
         CSoporte = 9999999
         FResistencia = ""
         CResistencia = 0
      
         'Para controlar la posicion en la matriz
         PosSoporte = 0
         PosResistencia = 0
      
         'Para controlar la posicion en la matriz
         PosSoporteUltimo = 0
         PosResistenciaUltima = 0
         
         ' Recorremos los volumenes
         For Y = 20 To 1 Step -1
         
             ' Si el volumen minimo del mercado es 0 o es mayor que el volumen del día
             If MatrizAcciones(Y, 6) < MatrizTemporal(Mercado, 5, 2) Or MatrizTemporal(Mercado, 5, 2) = 0 Then
             
                MatrizTemporal(Mercado, 5, 2) = MatrizAcciones(Y, 6)
             
             End If
             
             ' Si el volumen maximo del mercado es menor que el volumen del día
             If MatrizAcciones(Y, 6) > MatrizTemporal(Mercado, 5, 3) Then
             
                MatrizTemporal(Mercado, 5, 3) = MatrizAcciones(Y, 6)
             
             End If
              
             ' Agregamos a la medía de volumen el volumen del día
             MatrizTemporal(Mercado, 5, 1) = MatrizTemporal(Mercado, 5, 1) + MatrizAcciones(Y, 6)
              
             ' Si la vela minima del mercado es 0 o es mayor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) < MatrizTemporal(Mercado, 8, 2) Or MatrizTemporal(Mercado, 8, 2) = 0 Then
             
                MatrizTemporal(Mercado, 8, 2) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
             
             ' Si la vela maxima del mercado es menor que la vela del día
             If Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2))) > MatrizTemporal(Mercado, 8, 3) Then
             
                MatrizTemporal(Mercado, 8, 3) = Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
             
             End If
              
             ' Agregamos a la medía de vela la vela del día
             MatrizTemporal(Mercado, 8, 1) = MatrizTemporal(Mercado, 8, 1) + Abs((MatrizAcciones(Y, 3) - MatrizAcciones(Y, 2)))
              
             ' Si la cotización del cierre es menor o igual a la cotización del soporte del perido
             If CDbl(MatrizAcciones(Y, 3)) <= CDbl(CSoporte) Then
             
                FSoporte = MatrizAcciones(Y, 1)
                CSoporte = MatrizAcciones(Y, 3)
             
             End If
             
             ' Si la cotización del cierre es mayor o igual a la cotización de la resitencia del perido
             If CDbl(MatrizAcciones(Y, 3)) >= CDbl(CResistencia) Then
             
                FResistencia = MatrizAcciones(Y, 1)
                CResistencia = MatrizAcciones(Y, 3)
             
             End If
             
         Next
         
         ' Ajustamos el total de volumen entre total de dias
         MatrizTemporal(Mercado, 5, 1) = MatrizTemporal(Mercado, 5, 1) / 20
          
         ' Ajustamos el total de velas entre el total de días
         MatrizTemporal(Mercado, 8, 1) = MatrizTemporal(Mercado, 8, 1) / 20
         
         ' Actualizamos cotizaciones y fechas de soporte y resistencia
         MatrizTemporal(Mercado, 11, 1) = CSoporte
         MatrizTemporal(Mercado, 11, 2) = FSoporte
         MatrizTemporal(Mercado, 11, 3) = CResistencia
         MatrizTemporal(Mercado, 11, 4) = FResistencia
         
        'Si se ha marcado algun soporte
        If PosSoporte <> 0 Then
            
           DifSoporte = 9999999
         
           For Z = PosSoporte - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion del soporte dividido entre los dias es menor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z)) <= DifSoporte Then
                
                   DifSoporte = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CSoporte)) / (PosSoporte - Z))
                
                   PosSoporteUltimo = Z
                
                End If
            
            Next
            
            ' Actualizamos datos tendencia alcista 5 años
            MatrizTemporal(Mercado, 26, 1) = FSoporte
            MatrizTemporal(Mercado, 26, 2) = CSoporte
            
            If MatrizAcciones(PosSoporteUltimo, 1) = "" Then
                
               MatrizTemporal(Mercado, 26, 3) = FSoporte
            
            Else
               
               MatrizTemporal(Mercado, 26, 3) = MatrizAcciones(PosSoporteUltimo, 1)
               
            End If
            
            If MatrizAcciones(PosSoporteUltimo, 3) = "" Then
                
               MatrizTemporal(Mercado, 26, 4) = CSoporte
            
            Else
               
               MatrizTemporal(Mercado, 26, 4) = MatrizAcciones(PosSoporteUltimo, 3)
               
            End If
            
            If DifSoporte = 9999999 Then
                
               MatrizTemporal(Mercado, 26, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 26, 5) = Round(CDbl(DifSoporte) / (CDbl(CSoporte) * 0.01), 4)
               
            End If
            
            MatrizTemporal(Mercado, 26, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CSoporte)) / (CDbl(CSoporte) * 0.01), 4)
            MatrizTemporal(Mercado, 26, 7) = PosSoporte - 1
        
         End If
         
         'Si se ha marcado alguna resistencia
         If PosResistencia <> 0 Then
            
            DifResistencia = -9999999
         
            For Z = PosResistencia - 1 To 1 Step -1
            
                ' Si la cotizacion menos la cotizacion de la resistencia dividido entre los dias es mayor a la diferencia anterior
                If ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z)) >= DifResistencia Then
                
                   DifResistencia = ((CDbl(MatrizAcciones(Z, 3)) - CDbl(CResistencia)) / (PosResistencia - Z))
                
                   PosResistenciaUltima = Z
                   
                End If
            
            Next
            
            ' Actualizamos datos tendencia bajista 200 dias
            MatrizTemporal(Mercado, 27, 1) = FResistencia
            MatrizTemporal(Mercado, 27, 2) = CResistencia
            
            If MatrizAcciones(PosResistenciaUltima, 1) = "" Then
               
               MatrizTemporal(Mercado, 27, 3) = FResistencia
            
            Else
               
               MatrizTemporal(Mercado, 27, 3) = MatrizAcciones(PosResistenciaUltima, 1)
               
            End If
            
            If MatrizAcciones(PosResistenciaUltima, 3) = "" Then
                
               MatrizTemporal(Mercado, 27, 4) = CResistencia
            
            Else
               
               MatrizTemporal(Mercado, 27, 4) = MatrizAcciones(PosResistenciaUltima, 3)
               
            End If
            
            If DifResistencia = -9999999 Then
                
               MatrizTemporal(Mercado, 27, 5) = 0
            
            Else
               
               MatrizTemporal(Mercado, 27, 5) = Round(CDbl(DifResistencia) / (CDbl(CResistencia) * 0.01), 4)
              
            End If
            
            MatrizTemporal(Mercado, 27, 6) = Round((CDbl(MatrizAcciones(1, 3)) - CDbl(CResistencia)) / (CDbl(CResistencia) * 0.01), 4)
            MatrizTemporal(Mercado, 27, 7) = PosResistencia - 1
      
         End If
          
      End If
      
      ' PARA MARCAR LA TENDENCIA PREDOMINANTE
      
      ' Largo Plazo (5 años)
      
      ' Si los dias de actividad de ambas tendencias son distintos a 0
      If MatrizTemporal(Mercado, 14, 7) <> 0 And MatrizTemporal(Mercado, 15, 7) <> 0 Then
         
         ' Si los días activos de la alcista es superior a los dias de la bajista
         If MatrizTemporal(Mercado, 14, 7) > MatrizTemporal(Mercado, 15, 7) Then
      
            MatrizTemporal(Mercado, 14, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 14, 7) - MatrizTemporal(Mercado, 15, 7)) / MatrizTemporal(Mercado, 14, 7)))
            MatrizTemporal(Mercado, 15, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 14, 7) - MatrizTemporal(Mercado, 15, 7)) / MatrizTemporal(Mercado, 14, 7)))
            
         Else
      
            MatrizTemporal(Mercado, 14, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 14, 7) - MatrizTemporal(Mercado, 15, 7)) / MatrizTemporal(Mercado, 15, 7)))
            MatrizTemporal(Mercado, 15, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 14, 7) - MatrizTemporal(Mercado, 15, 7)) / MatrizTemporal(Mercado, 15, 7)))
      
         End If
         
      ' Si los dias de la tendencia alcista no son 0 y los de la bajista si
      ElseIf MatrizTemporal(Mercado, 14, 7) <> 0 And MatrizTemporal(Mercado, 15, 7) = 0 Then
      
         MatrizTemporal(Mercado, 14, 8) = AsignarTipoTendencia(1)
         MatrizTemporal(Mercado, 15, 8) = AsignarTipoTendencia(1)
         
      ' Si los dias de la tendencia alcista son 0 y los de la bajista no
      ElseIf MatrizTemporal(Mercado, 14, 7) = 0 And MatrizTemporal(Mercado, 15, 7) <> 0 Then
         
         MatrizTemporal(Mercado, 14, 8) = AsignarTipoTendencia(-1)
         MatrizTemporal(Mercado, 15, 8) = AsignarTipoTendencia(-1)
         
      End If
      
      ' Medio Plazo (200 dias)
      
      ' Si los dias de actividad de ambas tendencias son distintos a 0
      If MatrizTemporal(Mercado, 16, 7) <> 0 And MatrizTemporal(Mercado, 17, 7) <> 0 Then
         
         ' Si los días activos de la alcista es superior a los dias de la bajista
         If MatrizTemporal(Mercado, 16, 7) > MatrizTemporal(Mercado, 17, 7) Then
      
            MatrizTemporal(Mercado, 16, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 16, 7) - MatrizTemporal(Mercado, 17, 7)) / MatrizTemporal(Mercado, 16, 7)))
            MatrizTemporal(Mercado, 17, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 16, 7) - MatrizTemporal(Mercado, 17, 7)) / MatrizTemporal(Mercado, 16, 7)))
            
         Else
      
            MatrizTemporal(Mercado, 16, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 16, 7) - MatrizTemporal(Mercado, 17, 7)) / MatrizTemporal(Mercado, 17, 7)))
            MatrizTemporal(Mercado, 17, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 16, 7) - MatrizTemporal(Mercado, 17, 7)) / MatrizTemporal(Mercado, 17, 7)))
      
         End If
         
      ' Si los dias de la tendencia alcista no son 0 y los de la bajista si
      ElseIf MatrizTemporal(Mercado, 16, 7) <> 0 And MatrizTemporal(Mercado, 17, 7) = 0 Then
      
         MatrizTemporal(Mercado, 16, 8) = AsignarTipoTendencia(1)
         MatrizTemporal(Mercado, 17, 8) = AsignarTipoTendencia(1)
         
      ' Si los dias de la tendencia alcista son 0 y los de la bajista no
      ElseIf MatrizTemporal(Mercado, 16, 7) = 0 And MatrizTemporal(Mercado, 17, 7) <> 0 Then
         
         MatrizTemporal(Mercado, 16, 8) = AsignarTipoTendencia(-1)
         MatrizTemporal(Mercado, 17, 8) = AsignarTipoTendencia(-1)
         
      End If
      
      ' Corto Plazo (50 dias)
      
      ' Si los dias de actividad de ambas tendencias son distintos a 0
      If MatrizTemporal(Mercado, 18, 7) <> 0 And MatrizTemporal(Mercado, 19, 7) <> 0 Then
         
         ' Si los días activos de la alcista es superior a los dias de la bajista
         If MatrizTemporal(Mercado, 18, 7) > MatrizTemporal(Mercado, 19, 7) Then
      
            MatrizTemporal(Mercado, 18, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 18, 7) - MatrizTemporal(Mercado, 19, 7)) / MatrizTemporal(Mercado, 18, 7)))
            MatrizTemporal(Mercado, 19, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 18, 7) - MatrizTemporal(Mercado, 19, 7)) / MatrizTemporal(Mercado, 18, 7)))
            
         Else
      
            MatrizTemporal(Mercado, 18, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 18, 7) - MatrizTemporal(Mercado, 19, 7)) / MatrizTemporal(Mercado, 19, 7)))
            MatrizTemporal(Mercado, 19, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 18, 7) - MatrizTemporal(Mercado, 19, 7)) / MatrizTemporal(Mercado, 19, 7)))
      
         End If
         
      ' Si los dias de la tendencia alcista no son 0 y los de la bajista si
      ElseIf MatrizTemporal(Mercado, 18, 7) <> 0 And MatrizTemporal(Mercado, 19, 7) = 0 Then
      
         MatrizTemporal(Mercado, 18, 8) = AsignarTipoTendencia(1)
         MatrizTemporal(Mercado, 19, 8) = AsignarTipoTendencia(1)
         
      ' Si los dias de la tendencia alcista son 0 y los de la bajista no
      ElseIf MatrizTemporal(Mercado, 18, 7) = 0 And MatrizTemporal(Mercado, 19, 7) <> 0 Then
         
         MatrizTemporal(Mercado, 18, 8) = AsignarTipoTendencia(-1)
         MatrizTemporal(Mercado, 19, 8) = AsignarTipoTendencia(-1)
         
      End If
      
      ' Muy Corto Plazo (20 dias)
      
      ' Si los dias de actividad de ambas tendencias son distintos a 0
      If MatrizTemporal(Mercado, 26, 7) <> 0 And MatrizTemporal(Mercado, 27, 7) <> 0 Then
         
         ' Si los días activos de la alcista es superior a los dias de la bajista
         If MatrizTemporal(Mercado, 26, 7) > MatrizTemporal(Mercado, 27, 7) Then
      
            MatrizTemporal(Mercado, 26, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 26, 7) - MatrizTemporal(Mercado, 27, 7)) / MatrizTemporal(Mercado, 26, 7)))
            MatrizTemporal(Mercado, 27, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 26, 7) - MatrizTemporal(Mercado, 27, 7)) / MatrizTemporal(Mercado, 26, 7)))
            
         Else
      
            MatrizTemporal(Mercado, 26, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 26, 7) - MatrizTemporal(Mercado, 27, 7)) / MatrizTemporal(Mercado, 27, 7)))
            MatrizTemporal(Mercado, 27, 8) = AsignarTipoTendencia(((MatrizTemporal(Mercado, 26, 7) - MatrizTemporal(Mercado, 27, 7)) / MatrizTemporal(Mercado, 27, 7)))
      
         End If
         
      ' Si los dias de la tendencia alcista no son 0 y los de la bajista si
      ElseIf MatrizTemporal(Mercado, 26, 7) <> 0 And MatrizTemporal(Mercado, 27, 7) = 0 Then
      
         MatrizTemporal(Mercado, 26, 8) = AsignarTipoTendencia(1)
         MatrizTemporal(Mercado, 27, 8) = AsignarTipoTendencia(1)
         
      ' Si los dias de la tendencia alcista son 0 y los de la bajista no
      ElseIf MatrizTemporal(Mercado, 26, 7) = 0 And MatrizTemporal(Mercado, 27, 7) <> 0 Then
         
         MatrizTemporal(Mercado, 26, 8) = AsignarTipoTendencia(-1)
         MatrizTemporal(Mercado, 27, 8) = AsignarTipoTendencia(-1)
         
      End If
      
End Sub

Public Function AsignarTipoTendencia(Porcentaje As Double) As String

Porcentaje = Porcentaje * 100

If Porcentaje >= 60 Then

   AsignarTipoTendencia = "ALCISTA"

ElseIf Porcentaje <= -60 Then

   AsignarTipoTendencia = "BAJISTA"

ElseIf Porcentaje >= 40 And Porcentaje < 60 Then

   AsignarTipoTendencia = "LATERAL-ALCISTA"

ElseIf Porcentaje <= -40 And Porcentaje > -60 Then

   AsignarTipoTendencia = "LATERAL-BAJISTA"

Else

   AsignarTipoTendencia = "LATERAL"

End If

'+60% ALCISTA
'+40% LATERAL-ALCISTA
'     LATERAL
'-40% LATERAL-BAJISTA
'-60% BAJISTA

End Function

Public Sub AnalisisTecnicoMercados_MediasMoviles(Mercado As Integer)

      ' PARA MEDIAS MOVILES DE 200 DIAS n, n-1 ... n-4
      
      
      
      ' Recorremos desde 204 a 200 para calcular MM200, MM200n-1 ... MM200-4
      For Y = 207 To 200 Step -1
      
          ' Si el ultimo registro es númerico, entendemos que el resto tambien
          If IsNumeric(MatrizAcciones(Y, 3)) Then
      
             ' Recorremos los cierres desde el ultimo sumandolos y dividiendolos entre 200, para la media
             For Z = Y To Y - 199 Step -1
              
                 MatrizTemporal(Mercado, 4, ((Y - 200) + 2)) = MatrizTemporal(Mercado, 4, ((Y - 200) + 2)) + MatrizAcciones(Z, 3)
              
             Next
          
             MatrizTemporal(Mercado, 4, ((Y - 200) + 2)) = MatrizTemporal(Mercado, 4, ((Y - 200) + 2)) / 200
          
          End If
          
      Next
      
      ' Guardamos el signo de la ultima MM200 contra la penultima
      If MatrizTemporal(Mercado, 4, 2) = MatrizTemporal(Mercado, 4, 3) Then
      
         MatrizTemporal(Mercado, 4, 1) = "="
         
      ElseIf MatrizTemporal(Mercado, 4, 2) > MatrizTemporal(Mercado, 4, 3) Then
      
         MatrizTemporal(Mercado, 4, 1) = "+"
      
      Else
      
         MatrizTemporal(Mercado, 4, 1) = "-"
      
      End If
      
      ' PARA MEDIAS MOVILES DE 50 DIAS n, n-1 ... n-4
      
      'Recorremos desde 54 a 50 para calcular MM50, MM50n-1 ... MM50-4
      For Y = 57 To 50 Step -1
      
          'Si el ultimo registro es númerico, entendemos que el resto tambien
          If IsNumeric(MatrizAcciones(Y, 3)) Then
      
             'Recorremos los cierres desde el ultimo sumandolos y dividiendolos entre 50, para la media
             For Z = Y To Y - 49 Step -1
              
                 MatrizTemporal(Mercado, 3, ((Y - 50) + 2)) = MatrizTemporal(Mercado, 3, ((Y - 50) + 2)) + MatrizAcciones(Z, 3)
              
             Next
          
             MatrizTemporal(Mercado, 3, ((Y - 50) + 2)) = MatrizTemporal(Mercado, 3, ((Y - 50) + 2)) / 50
          
          End If
          
      Next
      
      'Guardamos el signo de la ultima MM50 contra la penultima
      If MatrizTemporal(Mercado, 3, 2) = MatrizTemporal(Mercado, 3, 3) Then
      
         MatrizTemporal(Mercado, 3, 1) = "="
         
      ElseIf MatrizTemporal(Mercado, 3, 2) > MatrizTemporal(Mercado, 3, 3) Then
      
         MatrizTemporal(Mercado, 3, 1) = "+"
      
      Else
      
         MatrizTemporal(Mercado, 3, 1) = "-"
      
      End If
      
      ' PARA MEDIAS MOVILES DE 20 DIAS n, n-1 ... n-4
      
      'Recorremos desde 24 a 20 para calcular MM20, MM20n-1 ... MM20-4
      For Y = 27 To 20 Step -1
      
          'Si el ultimo registro es númerico, entendemos que el resto tambien
          If IsNumeric(MatrizAcciones(Y, 3)) Then
      
             'Recorremos los cierres desde el ultimo sumandolos y dividiendolos entre 20, para la media
             For Z = Y To Y - 19 Step -1
              
                 MatrizTemporal(Mercado, 2, ((Y - 20) + 2)) = MatrizTemporal(Mercado, 2, ((Y - 20) + 2)) + MatrizAcciones(Z, 3)
              
             Next
          
             MatrizTemporal(Mercado, 2, ((Y - 20) + 2)) = MatrizTemporal(Mercado, 2, ((Y - 20) + 2)) / 20
          
          End If
          
      Next
      
      'Guardamos el signo de la ultima MM20 contra la penultima
      If MatrizTemporal(Mercado, 2, 2) = MatrizTemporal(Mercado, 2, 3) Then
      
         MatrizTemporal(Mercado, 2, 1) = "="
         
      ElseIf MatrizTemporal(Mercado, 2, 2) > MatrizTemporal(Mercado, 2, 3) Then
      
         MatrizTemporal(Mercado, 2, 1) = "+"
      
      Else
      
         MatrizTemporal(Mercado, 2, 1) = "-"
      
      End If

End Sub

Public Sub GuardarDefectosSQL()

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' PARA Guardar Lane
RutinaSQL = "UPDATE Defectos SET Lane_Periodo = [Lane_Periodo], Lane_K = [Lane_K], Lane_D = [Lane_D], Lane_DS = [Lane_DS], Lane_DSS = [Lane_DSS], Lane_SobreCompra = [Lane_SobreCompra], Lane_SobreVenta = [Lane_SobreVenta], Lane_AvisoClasica_Lento = [Lane_AvisoClasica_Lento], Lane_AvisoSZona_Lento = [Lane_AvisoSZona_Lento], Lane_AvisoPopCorn_Lento = [Lane_AvisoPopCorn_Lento], Lane_AvisoClasica_Rapido = [Lane_AvisoClasica_Rapido], Lane_AvisoSZona_Rapido = [Lane_AvisoSZona_Rapido], Lane_AvisoPopCorn_Rapido = [Lane_AvisoPopCorn_Rapido] WHERE Id_Defectos = '1'"

RutinaSQL = Replace(RutinaSQL, "[Lane_Periodo]", Lane_Periodo)
RutinaSQL = Replace(RutinaSQL, "[Lane_K]", Lane_K)
RutinaSQL = Replace(RutinaSQL, "[Lane_D]", Lane_D)
RutinaSQL = Replace(RutinaSQL, "[Lane_DS]", Lane_DS)
RutinaSQL = Replace(RutinaSQL, "[Lane_DSS]", Lane_DSS)
RutinaSQL = Replace(RutinaSQL, "[Lane_SobreCompra]", Lane_SobreCompra)
RutinaSQL = Replace(RutinaSQL, "[Lane_SobreVenta]", Lane_SobreVenta)
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoClasica_Lento]", CInt(Lane_AvisoClasica_Lento))
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoSZona_Lento]", CInt(Lane_AvisoSZona_Lento))
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoPopCorn_Lento]", CInt(Lane_AvisoPopCorn_Lento))
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoClasica_Rapido]", CInt(Lane_AvisoClasica_Rapido))
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoSZona_Rapido]", CInt(Lane_AvisoSZona_Rapido))
RutinaSQL = Replace(RutinaSQL, "[Lane_AvisoPopCorn_Rapido]", CInt(Lane_AvisoPopCorn_Rapido))

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Guardar RSI
RutinaSQL = "UPDATE Defectos SET RSI_Periodo = [RSI_Periodo], RSI_SobreCompra = [RSI_SobreCompra], RSI_SobreVenta = [RSI_SobreVenta], RSI_AvisoSalidaZona = [RSI_AvisoSalidaZona], RSI_AvisoFailureSwing = [RSI_AvisoFailureSwing], RSI_AvisoDivergencia = [RSI_AvisoDivergencia] WHERE Id_Defectos = '1'"

RutinaSQL = Replace(RutinaSQL, "[RSI_Periodo]", RSI_Periodo)
RutinaSQL = Replace(RutinaSQL, "[RSI_SobreCompra]", RSI_SobreCompra)
RutinaSQL = Replace(RutinaSQL, "[RSI_SobreVenta]", RSI_SobreVenta)
RutinaSQL = Replace(RutinaSQL, "[RSI_AvisoSalidaZona]", CInt(RSI_AvisoSalidaZona))
RutinaSQL = Replace(RutinaSQL, "[RSI_AvisoFailureSwing]", CInt(RSI_AvisoFailureSwing))
RutinaSQL = Replace(RutinaSQL, "[RSI_AvisoDivergencia]", CInt(RSI_AvisoDivergencia))

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Fecha obtención datos y fecha actualización
RutinaSQL = "UPDATE Defectos SET ActFecha = GETDATE() WHERE Id_Defectos = '1' "

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close

End Sub

Public Sub AnalisisTecnicoMercados_ActTabla(Mercado As Integer)

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

ExisteRegistro = False

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Id_Mercados FROM MercadosATecnico WHERE Id_Mercados = '" & CStr(MatrizTemporal(Mercado, 1, 1)) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
 
   ' Si nos devuelve un dato númerico es que existe el registro para el mercado
   If IsNumeric(RegistroSQL("Id_Mercados")) Then
     
      ExisteRegistro = True
   
   End If
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Si no existe registro en MercadosATecnico para este mercado
If ExisteRegistro = False Then

   RutinaSQL = "INSERT INTO MercadosATecnico (Id_Mercados) VALUES ('" & CStr(MatrizTemporal(Mercado, 1, 1)) & "')"

   Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End If

' PARA MM20
RutinaSQL = "UPDATE MercadosATecnico SET SignoMM20 = '[21]', MM20 = '[22]', MM20n1 = '[23]', MM20n2 = '[24]', MM20n3 = '[25]', MM20n4 = '[26]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 2, Y) = Replace(Round(MatrizTemporal(Mercado, 2, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[2" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 2, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA MM50
RutinaSQL = "UPDATE MercadosATecnico SET SignoMM50 = '[31]', MM50 = '[32]', MM50n1 = '[33]', MM50n2 = '[34]', MM50n3 = '[35]', MM50n4 = '[36]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 3, Y) = Replace(Round(MatrizTemporal(Mercado, 3, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[3" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 3, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA MM200
RutinaSQL = "UPDATE MercadosATecnico SET SignoMM200 = '[41]', MM200 = '[42]', MM200n1 = '[43]', MM200n2 = '[44]', MM200n3 = '[45]', MM200n4 = '[46]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 4, Y) = Replace(Round(MatrizTemporal(Mercado, 4, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[4" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 4, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol20
RutinaSQL = "UPDATE MercadosATecnico SET VolM20 = '[51]', VolMin20 = '[52]', VolMax20 = '[53]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 5, Y) = Replace(Round(MatrizTemporal(Mercado, 5, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[5" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 5, Y)))

Next


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol50
RutinaSQL = "UPDATE MercadosATecnico SET VolM50 = '[61]', VolMin50 = '[62]', VolMax50 = '[63]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 6, Y) = Replace(Round(MatrizTemporal(Mercado, 6, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[6" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 6, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol200
RutinaSQL = "UPDATE MercadosATecnico SET VolM200 = '[71]', VolMin200 = '[72]', VolMax200 = '[73]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 7, Y) = Replace(Round(MatrizTemporal(Mercado, 7, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[7" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 7, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela20
RutinaSQL = "UPDATE MercadosATecnico SET VelaM20 = '[81]', VelaMin20 = '[82]', VelaMax20 = '[83]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 8, Y) = Replace(Round(MatrizTemporal(Mercado, 8, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[8" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 8, Y)))

Next


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela50
RutinaSQL = "UPDATE MercadosATecnico SET VelaM50 = '[91]', VelaMin50 = '[92]', VelaMax50 = '[93]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 9, Y) = Replace(Round(MatrizTemporal(Mercado, 9, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[9" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 9, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela200
RutinaSQL = "UPDATE MercadosATecnico SET VelaM200 = '[101]', VelaMin200 = '[102]', VelaMax200 = '[103]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 10, Y) = Replace(Round(MatrizTemporal(Mercado, 10, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[10" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 10, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 20
RutinaSQL = "UPDATE MercadosATecnico SET Soporte20 = '[111]', FechaSoporte20 = CONVERT(DATETIME, '[112]', 103), Resistencia20 = '[113]', FechaResistencia20 = CONVERT(DATETIME, '[114]', 103) WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 11, Y) = Replace(Round(MatrizTemporal(Mercado, 11, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[11" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 11, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 50
RutinaSQL = "UPDATE MercadosATecnico SET Soporte50 = '[121]', FechaSoporte50 = CONVERT(DATETIME, '[122]', 103), Resistencia50 = '[123]', FechaResistencia50 = CONVERT(DATETIME, '[124]', 103) WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 12, Y) = Replace(Round(MatrizTemporal(Mercado, 12, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[12" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 12, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 200
RutinaSQL = "UPDATE MercadosATecnico SET Soporte200 = '[131]', FechaSoporte200 = CONVERT(DATETIME, '[132]', 103), Resistencia200 = '[133]', FechaResistencia200 = CONVERT(DATETIME, '[134]', 103) WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 13, Y) = Replace(Round(MatrizTemporal(Mercado, 13, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[13" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 13, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Largo
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TALargo = CONVERT(DATETIME, '[141]', 103), Valor1TALargo = '[142]', Fecha2TALargo = CONVERT(DATETIME, '[143]', 103), Valor2TALargo = '[144]', PorTALargo = '[145]', PorAcumuladoTALargo = '[146]', DiasTALargo = '[147]', TipoTALargo = '[148]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 14, Y) = Replace(Round(MatrizTemporal(Mercado, 14, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[14" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 14, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Largo
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TBLargo = CONVERT(DATETIME, '[151]', 103), Valor1TBLargo = '[152]', Fecha2TBLargo = CONVERT(DATETIME, '[153]', 103), Valor2TBLargo = '[154]', PorTBLargo = '[155]', PorAcumuladoTBLargo = '[156]', DiasTBLargo = '[157]', TipoTBLargo = '[158]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 15, Y) = Replace(Round(MatrizTemporal(Mercado, 15, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[15" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 15, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Medio
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TAMedio = CONVERT(DATETIME, '[161]', 103), Valor1TAMedio = '[162]', Fecha2TAMedio = CONVERT(DATETIME, '[163]', 103), Valor2TAMedio = '[164]', PorTAMedio = '[165]', PorAcumuladoTAMedio = '[166]', DiasTAMedio = '[167]', TipoTAMedio = '[168]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 16, Y) = Replace(Round(MatrizTemporal(Mercado, 16, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[16" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 16, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Medio
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TBMedio = CONVERT(DATETIME, '[171]', 103), Valor1TBMedio = '[172]', Fecha2TBMedio = CONVERT(DATETIME, '[173]', 103), Valor2TBMedio = '[174]', PorTBMedio = '[175]', PorAcumuladoTBMedio = '[176]', DiasTBMedio = '[177]', TipoTBMedio = '[178]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 17, Y) = Replace(Round(MatrizTemporal(Mercado, 17, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[17" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 17, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Corto
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TACorto = CONVERT(DATETIME, '[181]', 103), Valor1TACorto = '[182]', Fecha2TACorto = CONVERT(DATETIME, '[183]', 103), Valor2TACorto = '[184]', PorTACorto = '[185]', PorAcumuladoTACorto = '[186]', DiasTACorto = '[187]', TipoTACorto = '[188]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 18, Y) = Replace(Round(MatrizTemporal(Mercado, 18, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[18" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 18, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Corto
RutinaSQL = "UPDATE MercadosATecnico SET Fecha1TBCorto = CONVERT(DATETIME, '[191]', 103), Valor1TBCorto = '[192]', Fecha2TBCorto = CONVERT(DATETIME, '[193]', 103), Valor2TBCorto = '[194]', PorTBCorto = '[195]', PorAcumuladoTBCorto = '[196]', DiasTBCorto  = '[197]', TipoTBCorto = '[198]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 19, Y) = Replace(Round(MatrizTemporal(Mercado, 19, Y), 4), ",", ".")
    
    RutinaSQL = Replace(RutinaSQL, CStr("[19" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 19, Y)))

Next


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close

End Sub

Public Sub AnalisisTecnicoAcciones_ActTabla(Mercado As Integer)

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

ExisteRegistro = False

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Id_Acciones FROM AccionesATecnico WHERE Id_Acciones = '" & CStr(MatrizTemporal(Mercado, 1, 1)) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
 
   ' Si nos devuelve un dato númerico es que existe el registro para el mercado
   If IsNumeric(RegistroSQL("Id_Acciones")) Then
     
      ExisteRegistro = True
   
   End If
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Si no existe registro en MercadosATecnico para este mercado
If ExisteRegistro = False Then

   RutinaSQL = "INSERT INTO AccionesATecnico (Id_Acciones) VALUES ('" & CStr(MatrizTemporal(Mercado, 1, 1)) & "')"

   Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End If

' PARA MM20
RutinaSQL = "UPDATE AccionesATecnico SET SignoMM20 = '[21]', MM20 = '[22]', MM20n1 = '[23]', MM20n2 = '[24]', MM20n3 = '[25]', MM20n4 = '[26]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 2, Y) = Replace(Round(MatrizTemporal(Mercado, 2, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[2" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 2, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA MM50
RutinaSQL = "UPDATE AccionesATecnico SET SignoMM50 = '[31]', MM50 = '[32]', MM50n1 = '[33]', MM50n2 = '[34]', MM50n3 = '[35]', MM50n4 = '[36]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 3, Y) = Replace(Round(MatrizTemporal(Mercado, 3, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[3" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 3, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA MM200
RutinaSQL = "UPDATE AccionesATecnico SET SignoMM200 = '[41]', MM200 = '[42]', MM200n1 = '[43]', MM200n2 = '[44]', MM200n3 = '[45]', MM200n4 = '[46]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 6

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 Then MatrizTemporal(Mercado, 4, Y) = Replace(Round(MatrizTemporal(Mercado, 4, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[4" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 4, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol20
RutinaSQL = "UPDATE AccionesATecnico SET VolM20 = '[51]', VolMin20 = '[52]', VolMax20 = '[53]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 5, Y) = Replace(Round(MatrizTemporal(Mercado, 5, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[5" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 5, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol50
RutinaSQL = "UPDATE AccionesATecnico SET VolM50 = '[61]', VolMin50 = '[62]', VolMax50 = '[63]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 6, Y) = Replace(Round(MatrizTemporal(Mercado, 6, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[6" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 6, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Vol200
RutinaSQL = "UPDATE AccionesATecnico SET VolM200 = '[71]', VolMin200 = '[72]', VolMax200 = '[73]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 7, Y) = Replace(Round(MatrizTemporal(Mercado, 7, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[7" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 7, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela20
RutinaSQL = "UPDATE AccionesATecnico SET VelaM20 = '[81]', VelaMin20 = '[82]', VelaMax20 = '[83]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 8, Y) = Replace(Round(MatrizTemporal(Mercado, 8, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[8" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 8, Y)))

Next


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela50
RutinaSQL = "UPDATE AccionesATecnico SET VelaM50 = '[91]', VelaMin50 = '[92]', VelaMax50 = '[93]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 9, Y) = Replace(Round(MatrizTemporal(Mercado, 9, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[9" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 9, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Vela200
RutinaSQL = "UPDATE AccionesATecnico SET VelaM200 = '[101]', VelaMin200 = '[102]', VelaMax200 = '[103]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 3

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 10, Y) = Replace(Round(MatrizTemporal(Mercado, 10, Y), 4), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[10" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 10, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 20
RutinaSQL = "UPDATE AccionesATecnico SET Soporte20 = '[111]', FechaSoporte20 = CONVERT(DATETIME, '[112]', 103), Resistencia20 = '[113]', FechaResistencia20 = CONVERT(DATETIME, '[114]', 103) WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 11, Y) = Replace(Round(MatrizTemporal(Mercado, 11, Y), 4), ",", ".")

    If Y = 2 Or Y = 4 Then
    
       If MatrizTemporal(Mercado, 11, Y) = 0 Then MatrizTemporal(Mercado, 11, Y) = "1/1/1900"
    
    End If

    RutinaSQL = Replace(RutinaSQL, CStr("[11" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 11, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 50
RutinaSQL = "UPDATE AccionesATecnico SET Soporte50 = '[121]', FechaSoporte50 = CONVERT(DATETIME, '[122]', 103), Resistencia50 = '[123]', FechaResistencia50 = CONVERT(DATETIME, '[124]', 103) WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 12, Y) = Replace(Round(MatrizTemporal(Mercado, 12, Y), 4), ",", ".")
    
    If Y = 2 Or Y = 4 Then
    
       If MatrizTemporal(Mercado, 12, Y) = 0 Then MatrizTemporal(Mercado, 12, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[12" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 12, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Soportes y Resistencias 200
RutinaSQL = "UPDATE AccionesATecnico SET Soporte200 = '[131]', FechaSoporte200 = CONVERT(DATETIME, '[132]', 103), Resistencia200 = '[133]', FechaResistencia200 = CONVERT(DATETIME, '[134]', 103) WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 4

    ' Para cambiar coma por punto en campos númericos
    If Y = 1 Or Y = 3 Then MatrizTemporal(Mercado, 13, Y) = Replace(Round(MatrizTemporal(Mercado, 13, Y), 4), ",", ".")
    
    If Y = 2 Or Y = 4 Then
    
       If MatrizTemporal(Mercado, 13, Y) = 0 Then MatrizTemporal(Mercado, 13, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[13" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 13, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Largo
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TALargo = CONVERT(DATETIME, '[141]', 103), Valor1TALargo = '[142]', Fecha2TALargo = CONVERT(DATETIME, '[143]', 103), Valor2TALargo = '[144]', PorTALargo = '[145]', PorAcumuladoTALargo = '[146]', DiasTALargo = '[147]', TipoTALargo = '[148]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 14, Y) = Replace(Round(MatrizTemporal(Mercado, 14, Y), 4), ",", ".")
    
    If Y = 2 Or Y = 4 Then
    
       If MatrizTemporal(Mercado, 14, Y) = 0 Then MatrizTemporal(Mercado, 14, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[14" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 14, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Largo
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TBLargo = CONVERT(DATETIME, '[151]', 103), Valor1TBLargo = '[152]', Fecha2TBLargo = CONVERT(DATETIME, '[153]', 103), Valor2TBLargo = '[154]', PorTBLargo = '[155]', PorAcumuladoTBLargo = '[156]', DiasTBLargo = '[157]', TipoTBLargo = '[158]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 15, Y) = Replace(Round(MatrizTemporal(Mercado, 15, Y), 4), ",", ".")
    
    If Y = 2 Or Y = 4 Then
    
       If MatrizTemporal(Mercado, 15, Y) = 0 Then MatrizTemporal(Mercado, 15, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[15" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 15, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Medio
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TAMedio = CONVERT(DATETIME, '[161]', 103), Valor1TAMedio = '[162]', Fecha2TAMedio = CONVERT(DATETIME, '[163]', 103), Valor2TAMedio = '[164]', PorTAMedio = '[165]', PorAcumuladoTAMedio = '[166]', DiasTAMedio = '[167]', TipoTAMedio = '[168]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 16, Y) = Replace(Round(MatrizTemporal(Mercado, 16, Y), 4), ",", ".")
    
    If Y = 1 Or Y = 3 Then
    
       If MatrizTemporal(Mercado, 16, Y) = 0 Then MatrizTemporal(Mercado, 16, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[16" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 16, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Medio
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TBMedio = CONVERT(DATETIME, '[171]', 103), Valor1TBMedio = '[172]', Fecha2TBMedio = CONVERT(DATETIME, '[173]', 103), Valor2TBMedio = '[174]', PorTBMedio = '[175]', PorAcumuladoTBMedio = '[176]', DiasTBMedio = '[177]', TipoTBMedio = '[178]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 17, Y) = Replace(Round(MatrizTemporal(Mercado, 17, Y), 4), ",", ".")
    
    If Y = 1 Or Y = 3 Then
    
       If MatrizTemporal(Mercado, 17, Y) = 0 Then MatrizTemporal(Mercado, 17, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[17" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 17, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Alcista Corto
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TACorto = CONVERT(DATETIME, '[181]', 103), Valor1TACorto = '[182]', Fecha2TACorto = CONVERT(DATETIME, '[183]', 103), Valor2TACorto = '[184]', PorTACorto = '[185]', PorAcumuladoTACorto = '[186]', DiasTACorto = '[187]', TipoTACorto = '[188]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 18, Y) = Replace(Round(MatrizTemporal(Mercado, 18, Y), 4), ",", ".")
    
    If Y = 1 Or Y = 3 Then
    
       If MatrizTemporal(Mercado, 18, Y) = 0 Then MatrizTemporal(Mercado, 18, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[18" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 18, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Tendencia Bajista Corto
RutinaSQL = "UPDATE AccionesATecnico SET Fecha1TBCorto = CONVERT(DATETIME, '[191]', 103), Valor1TBCorto = '[192]', Fecha2TBCorto = CONVERT(DATETIME, '[193]', 103), Valor2TBCorto = '[194]', PorTBCorto = '[195]', PorAcumuladoTBCorto = '[196]', DiasTBCorto  = '[197]', TipoTBCorto = '[198]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 8

    ' Para cambiar coma por punto en campos númericos
    If Y <> 1 And Y <> 3 And Y <> 8 Then MatrizTemporal(Mercado, 19, Y) = Replace(Round(MatrizTemporal(Mercado, 19, Y), 4), ",", ".")
    
    If Y = 1 Or Y = 3 Then
    
       If MatrizTemporal(Mercado, 19, Y) = 0 Then MatrizTemporal(Mercado, 19, Y) = "1/1/1900"
    
    End If
    
    RutinaSQL = Replace(RutinaSQL, CStr("[19" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 19, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)
' Cerramos los objetos abiertos para la conexión

' PARA Fecha obtención datos y fecha actualización
RutinaSQL = "UPDATE AccionesATecnico SET FechaDatos = CONVERT(DATETIME, '[12]', 103), ActFecha = GETDATE() WHERE Id_Acciones = '[11]' "

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[12]", CStr(MatrizTemporal(Mercado, 1, 2)))


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close

End Sub

Public Sub AnalisisTecnicoMercados_ActTabla_Indicadores(Mercado As Integer)

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

ExisteRegistro = False

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Id_Mercados FROM Mercados_Indicadores WHERE Id_Mercados = '" & CStr(MatrizTemporal(Mercado, 1, 1)) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
 
   ' Si nos devuelve un dato númerico es que existe el registro para el mercado
   If IsNumeric(RegistroSQL("Id_Mercados")) Then
     
      ExisteRegistro = True
   
   End If
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Si no existe registro en MercadosATecnico para este mercado
If ExisteRegistro = False Then

   RutinaSQL = "INSERT INTO Mercados_Indicadores (Id_Mercados) VALUES ('" & CStr(MatrizTemporal(Mercado, 1, 1)) & "')"

   Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End If

' PARA K
RutinaSQL = "UPDATE Mercados_Indicadores SET K = '[201]', Kn1 = '[202]', Kn2 = '[203]', Kn3 = '[204]', Kn4 = '[205]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 20, Y) = Replace(MatrizTemporal(Mercado, 20, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[20" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 20, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA D
RutinaSQL = "UPDATE Mercados_Indicadores SET D = '[211]', Dn1 = '[212]', Dn2 = '[213]', Dn3 = '[214]', Dn4 = '[215]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 21, Y) = Replace(MatrizTemporal(Mercado, 21, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[21" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 21, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA DS
RutinaSQL = "UPDATE Mercados_Indicadores SET DS = '[221]', DSn1 = '[222]', DSn2 = '[223]', DSn3 = '[224]', DSn4 = '[225]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 22, Y) = Replace(MatrizTemporal(Mercado, 22, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[22" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 22, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA DSS
RutinaSQL = "UPDATE Mercados_Indicadores SET DSS = '[231]', DSSn1 = '[232]', DSSn2 = '[233]', DSSn3 = '[234]', DSSn4 = '[235]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 23, Y) = Replace(MatrizTemporal(Mercado, 23, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[23" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 23, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA RSI
RutinaSQL = "UPDATE Mercados_Indicadores SET RSI = '[241]', RSIn1 = '[242]', RSIn2 = '[243]', RSIn3 = '[244]', RSIn4 = '[245]' WHERE Id_Mercados = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 24, Y) = Replace(MatrizTemporal(Mercado, 24, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[24" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 24, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Fecha obtención datos y fecha actualización
RutinaSQL = "UPDATE Mercados_Indicadores SET FechaDatos = CONVERT(DATETIME, '[12]', 103), ActFecha = GETDATE() WHERE Id_Mercados = '[11]' "

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[12]", CStr(MatrizTemporal(Mercado, 1, 2)))


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close

End Sub

Public Sub AnalisisTecnicoAcciones_ActTabla_Indicadores(Mercado As Integer)

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

ExisteRegistro = False

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Id_Acciones FROM Acciones_Indicadores WHERE Id_Acciones = '" & CStr(MatrizTemporal(Mercado, 1, 1)) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
 
   ' Si nos devuelve un dato númerico es que existe el registro para el mercado
   If IsNumeric(RegistroSQL("Id_Acciones")) Then
     
      ExisteRegistro = True
   
   End If
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Si no existe registro en MercadosATecnico para este mercado
If ExisteRegistro = False Then

   RutinaSQL = "INSERT INTO Acciones_Indicadores (Id_Acciones) VALUES ('" & CStr(MatrizTemporal(Mercado, 1, 1)) & "')"

   Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End If

' PARA K
RutinaSQL = "UPDATE Acciones_Indicadores SET K = '[201]', Kn1 = '[202]', Kn2 = '[203]', Kn3 = '[204]', Kn4 = '[205]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 20, Y) = Replace(MatrizTemporal(Mercado, 20, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[20" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 20, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA D
RutinaSQL = "UPDATE Acciones_Indicadores SET D = '[211]', Dn1 = '[212]', Dn2 = '[213]', Dn3 = '[214]', Dn4 = '[215]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 21, Y) = Replace(MatrizTemporal(Mercado, 21, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[21" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 21, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA DS
RutinaSQL = "UPDATE Acciones_Indicadores SET DS = '[221]', DSn1 = '[222]', DSn2 = '[223]', DSn3 = '[224]', DSn4 = '[225]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 22, Y) = Replace(MatrizTemporal(Mercado, 22, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[22" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 22, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA DSS
RutinaSQL = "UPDATE Acciones_Indicadores SET DSS = '[231]', DSSn1 = '[232]', DSSn2 = '[233]', DSSn3 = '[234]', DSSn4 = '[235]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 23, Y) = Replace(MatrizTemporal(Mercado, 23, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[23" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 23, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA RSI
RutinaSQL = "UPDATE Acciones_Indicadores SET RSI = '[241]', RSIn1 = '[242]', RSIn2 = '[243]', RSIn3 = '[244]', RSIn4 = '[245]' WHERE Id_Acciones = '[11]'"

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

For Y = 1 To 5

    ' Para cambiar coma por punto en campos númericos
    MatrizTemporal(Mercado, 24, Y) = Replace(MatrizTemporal(Mercado, 24, Y), ",", ".")

    RutinaSQL = Replace(RutinaSQL, CStr("[24" & CStr(Y) & "]"), CStr(MatrizTemporal(Mercado, 24, Y)))

Next

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' PARA Fecha obtención datos y fecha actualización
RutinaSQL = "UPDATE Acciones_Indicadores SET FechaDatos = CONVERT(DATETIME, '[12]', 103), ActFecha = GETDATE() WHERE Id_Acciones = '[11]' "

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[11]", CStr(MatrizTemporal(Mercado, 1, 1)))

' Filtramos el Id del mercado
RutinaSQL = Replace(RutinaSQL, "[12]", CStr(MatrizTemporal(Mercado, 1, 2)))


Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close

End Sub
Public Sub InicializarAplicacion()

FicheroUDLSQL = "FILE NAME=" & App.Path & "\ConexionSQL.udl"

RutaRegistro = App.Path & "\Log\" & Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2) & "-Registro.txt"
RutaAvisosMercados = App.Path & "\Log\" & Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2) & "-AvisosMercados.txt"
RutaAvisosAcciones = App.Path & "\Log\" & Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2) & "-AvisosAcciones.txt"

CargarDefectos

CargarTiposVelas

End Sub

Sub GuardarAvisosMercados(Mercado As Integer)
  
If MatrizAvisos(Mercado, 0, 0) <> 0 Then

   ' Abrimos el fichero de registro en modo agregar
   Open RutaAvisosMercados For Append As #1
    
        ' Escribimos una línea de separación en el registro
        Print #1, "------------------------------------------------------------"
   
        ' Escribimos el nombre del mercado, la fecha de la ultima cotización y la fecha y hora de generacion
        Print #1, MatrizTemporal(Mercado, 1, 4) & " - " & MatrizTemporal(Mercado, 1, 2) & " - " & Now
   
        For Y = 1 To CInt(MatrizAvisos(Mercado, 0, 0))
        
            ' Imprimimos una línea en blanco
            Print #1, " "

            ' Escribimos la orden y donde se ha producido
            Print #1, " " & MatrizAvisos(Mercado, Y, 1) & " - " & MatrizAvisos(Mercado, Y, 2)
            
            ' Escribimos los controles considerables
            Print #1, " " & MatrizAvisos(Mercado, Y, 3)
            
            ' Si en confirmaciones hay algo
            If Len(MatrizAvisos(Mercado, Y, 4)) <> 0 Then
            
               ' Escribimos las confirmaciones
               Print #1, " " & MatrizAvisos(Mercado, Y, 4)
            
            End If
            
            ' Si en procedimientos hay algo
            If Len(MatrizAvisos(Mercado, Y, 5)) <> 0 Then
            
               ' Escribimos las confirmaciones
               Print #1, " " & MatrizAvisos(Mercado, Y, 5)
            
            End If
           
         Next
    
   ' Cerramos el fichero de registro
   Close #1
   
End If

End Sub

Sub GuardarAvisosAcciones(Mercado As Integer)
  
If MatrizAvisos(Mercado, 0, 0) <> 0 Then

   ' Abrimos el fichero de registro en modo agregar
   Open RutaAvisosAcciones For Append As #1
    
        ' Escribimos una línea de separación en el registro
        Print #1, "------------------------------------------------------------"
   
        ' Escribimos el nombre del mercado, la fecha de la ultima cotización y la fecha y hora de generacion
        Print #1, MatrizTemporal(Mercado, 1, 4) & " - " & MatrizTemporal(Mercado, 1, 2) & " - " & Now
   
        For Y = 1 To CInt(MatrizAvisos(Mercado, 0, 0))
        
            ' Imprimimos una línea en blanco
            Print #1, " "

            ' Escribimos la orden y donde se ha producido
            Print #1, " " & MatrizAvisos(Mercado, Y, 1) & " - " & MatrizAvisos(Mercado, Y, 2)
            
            ' Escribimos los controles considerables
            Print #1, " " & MatrizAvisos(Mercado, Y, 3)
            
            ' Si en confirmaciones hay algo
            If Len(MatrizAvisos(Mercado, Y, 4)) <> 0 Then
            
               ' Escribimos las confirmaciones
               Print #1, " " & MatrizAvisos(Mercado, Y, 4)
            
            End If
            
            ' Si en procedimientos hay algo
            If Len(MatrizAvisos(Mercado, Y, 5)) <> 0 Then
            
               ' Escribimos las confirmaciones
               Print #1, " " & MatrizAvisos(Mercado, Y, 5)
            
            End If
           
         Next
    
   ' Cerramos el fichero de registro
   Close #1
   
End If

End Sub

Function EscribirRegistro(RutaFichero As String, Mensaje As String) As Integer
  
' Abrimos el fichero de registro en modo agregar
Open RutaFichero For Append As #1
      
    ' Escribimos el mensaje en le fichero de registro
    Print #1, Mensaje
    
' Cerramos el fichero de registro
Close #1

End Function

Public Function ExisteSQL() As String

On Error GoTo ErrorExisteSQL

' Por si estan abiertas
Set ConexionSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

ConexionSQL.Close
        
ErrorExisteSQL:

    If InStr(1, Err.Description, "No existe el servidor SQL Server", vbTextCompare) <> 0 Then
         
       ExisteSQL = "No existe el servidor SQL Server o se ha denegado el acceso al mismo."
    
    ElseIf InStr(1, Err.Description, "Cannot open database requested in login", vbTextCompare) <> 0 Then
    
       ExisteSQL = "No existe la base de datos indicada en el servidor SQL Server."
    
    ElseIf InStr(1, Err.Description, "Login failed for user", vbTextCompare) <> 0 Then
    
       ExisteSQL = "Servidor SQL Server NO accesible con usuario y contraseña inidicados"
    
    Else
      
       ExisteSQL = Err.Description
    
    End If
       
    Exit Function

End Function

' DESCARGA DE UN FICHERO DE INTERNET A UNA RUTA LOCAL
' Indicamos el nombre del fichero URL y la ruta donde queremos almacenarlo
' Dependencias (URLDownloadToFile de "urlmon")

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean

   Dim lngRetVal As Long
   
   lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
   
   If lngRetVal = 0 Then DownloadFile = True

End Function

' LECTURA DE LA TABLA MERCADOS SQL, RECOGIENDO TICKER DE LOS ACTIVADOS
'
' Leemos los TickerYahoo de la tabla Mercados, de aquellos registros que tengan
' activa la opción de ControlValores

Public Sub CargarMercados()

VarLinMatrizMercados = 0

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Count(Id_Mercados) As NumeroRegistros FROM Mercados", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim Preserve MatrizMercados(RegistroSQL("NumeroRegistros"), 6)
      
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
    
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT * FROM Mercados", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   
   VarLinMatrizMercados = VarLinMatrizMercados + 1
      
   MatrizMercados(VarLinMatrizMercados, 0) = RegistroSQL("TickerYahoo")
   MatrizMercados(VarLinMatrizMercados, 1) = RegistroSQL("Nombre")
   MatrizMercados(VarLinMatrizMercados, 2) = RegistroSQL("Zona")
   MatrizMercados(VarLinMatrizMercados, 3) = CBool(RegistroSQL("ControlHis"))
   MatrizMercados(VarLinMatrizMercados, 4) = CBool(RegistroSQL("ControlCot"))
   MatrizMercados(VarLinMatrizMercados, 5) = CBool(RegistroSQL("ControlValorHis"))
   MatrizMercados(VarLinMatrizMercados, 6) = CBool(RegistroSQL("ControlValorCot"))
      
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub

Sub AsignarValorCampoCotizacion(NumeroCampo As Integer, ValorCampo As String)

Select Case NumeroCampo
       Case 1
         CotFecha = ValorCampo
       Case 2
         CotApertura = ValorCampo
       Case 3
         CotMaximo = ValorCampo
       Case 4
         CotMinimo = ValorCampo
       Case 5
         CotCierre = ValorCampo
       Case 6
         CotVolumen = ValorCampo
End Select

End Sub

Public Sub CargarAccionesImportacionCotizaciones(Origen As String)

TotalProcedimiento = 5
UnidadesProcedimiento = 0

VarLinMatrizAcciones = 0

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset
    
' Abrimos la conexión con la base de datos de SQL
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT COUNT(Acciones.Id_Acciones) As NumeroRegistros FROM Acciones INNER JOIN AccionesMercados ON AccionesMercados.Id_Acciones = Acciones.Id_Acciones INNER JOIN Mercados ON Mercados.Id_Mercados = AccionesMercados.Id_Mercados WHERE Mercados.ControlValorHis = 1 AND IsNull(Acciones.Ticker" & Origen & ", '') <> ''", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim MatrizAcciones(RegistroSQL("NumeroRegistros"), 3)
      
   If CLng(RegistroSQL("NumeroRegistros")) <> 0 Then
   
      UnidadesTotalesProcedimiento = CLng(RegistroSQL("NumeroRegistros"))
   
      UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento
      
   End If
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
    
' Abrimos el recordset de la tabla Acciones, con los datos necesarios para
RegistroSQL.Open "SELECT Acciones.Id_Acciones, IsNull(Acciones.Ticker" & Origen & ", '') As Ticker" & Origen & ", MAX(DATEADD(DAY, 1, IsNull(AccionesCotizaciones.Fecha, '1/1/1900'))) As FechaDesde FROM Acciones INNER JOIN AccionesMercados ON AccionesMercados.Id_Acciones = Acciones.Id_Acciones INNER JOIN Mercados ON Mercados.Id_Mercados = AccionesMercados.Id_Mercados LEFT OUTER JOIN AccionesCotizaciones ON AccionesCotizaciones.Id_Acciones = Acciones.Id_Acciones WHERE Mercados.ControlValorHis = 1 AND IsNull(AccionesCotizaciones.Origen, 'I') = 'I' AND IsNull(Acciones.Ticker" & Origen & ", '') <> '' GROUP BY Mercados.Id_Mercados, Acciones.Id_Acciones, Acciones.Ticker" & Origen & " ORDER BY Mercados.Id_Mercados, Acciones.Ticker" & Origen, ConexionSQL, adOpenDynamic, adLockOptimistic

' Mientras que haya registros
Do While Not RegistroSQL.EOF

   UnidadesProcedimiento = UnidadesProcedimiento + 1
   
   FrmFiltroImpCotAcciones.BarraProgreso.Value = FrmFiltroImpCotAcciones.BarraProgreso.Value + UnitarioProcedimiento
   FrmFiltroImpCotAcciones.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Cargando " & RegistroSQL("Ticker" & Origen)
   
   DoEvents
   
   ' Agregamos 1 al contador de filas de la matriz
   VarLinMatrizAcciones = VarLinMatrizAcciones + 1
      
   ' Introducimos en las columnas el identificador de la acción y el ticker según origen seleccionado (TickerYahoo)
   MatrizAcciones(VarLinMatrizAcciones, 0) = RegistroSQL("Id_Acciones")
   MatrizAcciones(VarLinMatrizAcciones, 1) = RegistroSQL("Ticker" & Origen)
   
   ' Si la fecha + 1 día, del último registro de cotizaciones de la acción es mayor que la fecha desde solicitada
   If CDate(RegistroSQL("FechaDesde")) > CDate(VarFechaDesde) Then
   
      ' Ponemos la fecha del registro como fecha desde
      MatrizAcciones(VarLinMatrizAcciones, 2) = RegistroSQL("FechaDesde")
      
   Else
      
      ' Mantenemos la fecha desde solicitada
      MatrizAcciones(VarLinMatrizAcciones, 2) = VarFechaDesde
      
   End If
   
   ' Introducimos la fecha hasta solicitada
   MatrizAcciones(VarLinMatrizAcciones, 3) = VarFechaHasta

   ' Leemos el siguiente registro
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub

Public Sub DescargarAccionesImportacionCotizaciones(Origen As String)

TotalProcedimiento = 15
UnidadesProcedimiento = 0

If VarLinMatrizAcciones <> 0 Then

   UnidadesTotalesProcedimiento = VarLinMatrizAcciones

   UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento
   
End If

' Recorremos las líneas de MatrizMercados
For i = 1 To VarLinMatrizAcciones
    
    UnidadesProcedimiento = UnidadesProcedimiento + 1

    FrmFiltroImpCotAcciones.BarraProgreso.Value = FrmFiltroImpCotAcciones.BarraProgreso.Value + UnitarioProcedimiento
    FrmFiltroImpCotAcciones.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Descargando CSV cotizaciones " & MatrizAcciones(i, 1)
   
    DoEvents
    
    ' Si el origen de datos solicitado es Yahoo
    If Origen = "Yahoo" Then
       
       ' Montamos una cadena de texto con la base de la URL de petición de datos de acciones a Yahoo
       VarLinkYahoo = "http://ichart.yahoo.com/table.csv?s=[TickerYahoo]&d=[MesHasta]&e=[DiaHasta]&f=[AñoHasta]&g=d&a=[MesDesde]&b=[DiaDesde]&c=[AñoDesde]&ignore=.csv"

       ' Sustituimos de la cadena de texto base la etiqueta del Ticker por el ticker real de la acción guardado en la matriz
       VarLinkYahoo = Replace(VarLinkYahoo, "[TickerYahoo]", CStr(MatrizAcciones(i, 1)))
       
       ' Sustituimos de la cadena de texto base las etiquetas de datos de fecha desde, por los datos de la matriz
       VarLinkYahoo = Replace(VarLinkYahoo, "[MesHasta]", CStr(Month(MatrizAcciones(i, 3))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaHasta]", CStr(Day(MatrizAcciones(i, 3))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoHasta]", CStr(Year(MatrizAcciones(i, 3))))
       
       ' Sustituimos de la cadena de texto base las etiquetas de datos de fecha hasta, por los datos de la matriz
       VarLinkYahoo = Replace(VarLinkYahoo, "[MesDesde]", CStr(Month(MatrizAcciones(i, 2))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaDesde]", CStr(Day(MatrizAcciones(i, 2))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoDesde]", CStr(Year(MatrizAcciones(i, 2))))

       ' Asignamos a la variable la URL montada
       VarNombreURL = VarLinkYahoo
       ' Asignamos como ruta de destino local, la carpteta Temp dentro de la ruta del programa y como nombre de fichero el Ticker con extensión CSV
       VarNombreFichero = CurDir & "\Temp\" & MatrizAcciones(i, 1) & ".csv"
    
    ElseIf Origen = "Economista" Then
    
    'http://www.eleconomista.es/descargas-historicos/empresa/MAPFRE/2016-08-1/2016-08-28
    
       VarLinkYahoo = "http://www.eleconomista.es/descargas-historicos/empresa/[TickerYahoo]/[AñoDesde]-[MesDesde]-[DiaDesde]/[AñoHasta]-[MesHasta]-[DiaHasta]"
       'VarLinkYahoo = "http://ichart.yahoo.com/table.csv?s=[TickerYahoo]&d=[MesHasta]&e=[DiaHasta]&f=[AñoHasta]&g=d&a=[MesDesde]&b=[DiaDesde]&c=[AñoDesde]&ignore=.csv"

       VarLinkYahoo = Replace(VarLinkYahoo, "[TickerYahoo]", CStr(MatrizAcciones(i, 1)))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesHasta]", CStr(Month(MatrizAcciones(i, 3))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaHasta]", CStr(Day(MatrizAcciones(i, 3))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoHasta]", CStr(Year(MatrizAcciones(i, 3))))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesDesde]", CStr(Month(MatrizAcciones(i, 2))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaDesde]", CStr(Day(MatrizAcciones(i, 2))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoDesde]", CStr(Year(MatrizAcciones(i, 2))))

       '[TickerYahoo] TL5.MC, [MesHasta] 3, [DiaHasta] 15, [AñoHasta] 2016, [MesDesde] 5, [DiaDesde] 24, [AñoDesde] 2015
    
       VarNombreURL = VarLinkYahoo
       VarNombreFichero = CurDir & "\Temp\" & MatrizAcciones(i, 1) & ".csv"
    
    ElseIf Origen = "Invertia" Then
    
    'http://www.invertia.com/inc/bolsa/ficha/excel.asp?FechaDesde=2015/11/01%2000:00&FechaHasta=2015/11/25%2000:00&idtel=RV011ABENGOA
    
       VarLinkYahoo = "http://www.invertia.com/inc/bolsa/ficha/excel.asp?FechaDesde=[AñoDesde]/[MesDesde]/[DiaDesde]%2000:00&FechaHasta=[AñoHasta]/[MesHasta]/[DiaHasta]%2000:00&idtel=RV011[TickerYahoo]"
       'VarLinkYahoo = "http://www.eleconomista.es/descargas-historicos/empresa/[TickerYahoo]/[AñoDesde]-[MesDesde]-[DiaDesde]/[AñoHasta]-[MesHasta]-[DiaHasta]"

       VarLinkYahoo = Replace(VarLinkYahoo, "[TickerYahoo]", CStr(MatrizAcciones(i, 1)))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesHasta]", CStr(Month(MatrizAcciones(i, 3))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaHasta]", CStr(Day(MatrizAcciones(i, 3))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoHasta]", CStr(Year(MatrizAcciones(i, 3))))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesDesde]", CStr(Month(MatrizAcciones(i, 2))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaDesde]", CStr(Day(MatrizAcciones(i, 2))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoDesde]", CStr(Year(MatrizAcciones(i, 2))))

       '[TickerYahoo] TL5.MC, [MesHasta] 3, [DiaHasta] 15, [AñoHasta] 2016, [MesDesde] 5, [DiaDesde] 24, [AñoDesde] 2015
    
       VarNombreURL = VarLinkYahoo
       VarNombreFichero = CurDir & "\Temp\" & MatrizAcciones(i, 1) & ".xls"
       
    End If
    
    ' Llamamos al procedimiento Download File para descargar con los parametros URL petición y ruta descarga
    DownloadFile VarNombreURL, VarNombreFichero

    'MsgBox VarNombreURL & " | " & VarNombreFichero
        
Next

End Sub

Sub ImportarCotizacionesAcciones(IdAcciones As String, Fecha As String, Apertura As String, Cierre As String, Maximo As String, Minimo As String, Volumen As String, Origen As String)

' Por si estan abiertas
Set RegistroSQL = Nothing

' Creamos los objetos
Set RegistroSQL = New ADODB.Recordset

' Si el origen es cotización actual
If Origen = "C" Then

   ' Abrimos el recordset de la consulta
   RegistroSQL.Open "SELECT Id_Acciones FROM Acciones WHERE TickerYahoo = '" & IdAcciones & "'", ConexionSQL, adOpenDynamic, adLockOptimistic
   
   Do While Not RegistroSQL.EOF

      'MsgBox RegistroSQL("Id_Acciones")
   
      IdAcciones = CStr(RegistroSQL("Id_Acciones"))
      
      RegistroSQL.MoveNext
    
   Loop

   ' Cerramos los objetos abiertos para la conexión
   RegistroSQL.Close

End If

' Para filtrar que no sea una importación cotización de una accion que no existe
If IsNumeric(IdAcciones) Then

   ' Abrimos el recordset de la consulta
   RegistroSQL.Open "SELECT Id_Acciones FROM AccionesCotizaciones WHERE Id_Acciones = '" & CStr(IdAcciones) & "' AND Fecha = '" & CStr(Year(Fecha) & Right("00" & Month(Fecha), 2) & Right("00" & Day(Fecha), 2)) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

   ' Montamos la rutina para insertar
   RutinaSQL = "INSERT INTO AccionesCotizaciones (Id_Acciones, Fecha, Apertura, Cierre, Maximo, Minimo, Volumen, Origen, ActFecha) VALUES ('[IdAcciones]', '[Fecha]', '[Apertura]', '[Cierre]', '[Maximo]', '[Minimo]', '[Volumen]', '[Origen]', GETDATE())"

   Do While Not RegistroSQL.EOF

      'MsgBox RegistroSQL("Id_Acciones")
   
      ' Montamos la rutina para actualizar porque ya existe el registro
      RutinaSQL = "UPDATE AccionesCotizaciones SET Apertura = '[Apertura]', Cierre = '[Cierre]', Maximo = '[Maximo]', Minimo = '[Minimo]', Volumen = '[Volumen]', Origen = '[Origen]', ActFecha = GETDATE() WHERE Id_Acciones = '[IdAcciones]' AND Fecha = '[Fecha]'"
   
      RegistroSQL.MoveNext
    
   Loop

   ' Cerramos los objetos abiertos para la conexión
   RegistroSQL.Close
 
   ' Sustituimos las etiquetas entre corchetes, por los datos recogidos de los ficheros de cotizaciones CSV
   RutinaSQL = Replace(RutinaSQL, "[IdAcciones]", CStr(IdAcciones))
   RutinaSQL = Replace(RutinaSQL, "[Fecha]", CStr(Year(Fecha) & Right("00" & Month(Fecha), 2) & Right("00" & Day(Fecha), 2)))
   RutinaSQL = Replace(RutinaSQL, "[Apertura]", CStr(Replace(Apertura, ",", ".")))
   RutinaSQL = Replace(RutinaSQL, "[Cierre]", CStr(Replace(Cierre, ",", ".")))
   RutinaSQL = Replace(RutinaSQL, "[Maximo]", CStr(Replace(Maximo, ",", ".")))
   RutinaSQL = Replace(RutinaSQL, "[Minimo]", CStr(Replace(Minimo, ",", ".")))
   RutinaSQL = Replace(RutinaSQL, "[Volumen]", CStr(Replace(Volumen, ",", ".")))
   RutinaSQL = Replace(RutinaSQL, "[Origen]", CStr(Origen))

   ' Ejecutamos la rutina de inserción de datos en la tabla AccionesCotizaciones
   Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End If

End Sub

Public Sub CargarMercadosImportacionCotizaciones(Origen As String)

TotalProcedimiento = 5
UnidadesProcedimiento = 0

VarLinMatrizAcciones = 0

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT COUNT(Mercados.Id_Mercados) As NumeroRegistros FROM Mercados WHERE Mercados.ControlHis = 1", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF
   
   'Recogemos el número de registros de la tabla Mercados y redimensionamos la tabla donde van
   ReDim MatrizAcciones(RegistroSQL("NumeroRegistros"), 3)
      
   UnidadesTotalesProcedimiento = CLng(RegistroSQL("NumeroRegistros"))
   
   UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento
   
   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
    
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Mercados.Id_Mercados, Mercados.Ticker" & Origen & ", MAX(DATEADD(DAY, 1, IsNull(MercadosCotizaciones.Fecha, '1/1/1900'))) As FechaDesde FROM Mercados LEFT OUTER JOIN MercadosCotizaciones ON MercadosCotizaciones.Id_Mercados = Mercados.Id_Mercados WHERE Mercados.ControlHis = 1 AND IsNull(MercadosCotizaciones.Origen, 'I') = 'I' GROUP BY Mercados.Id_Mercados, Mercados.Ticker" & Origen & " ORDER BY Mercados.Id_Mercados, Mercados.Ticker" & Origen, ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   UnidadesProcedimiento = UnidadesProcedimiento + 1
   
   FrmFiltroImpCotMercados.BarraProgreso.Value = FrmFiltroImpCotMercados.BarraProgreso.Value + UnitarioProcedimiento
   FrmFiltroImpCotMercados.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Cargando " & RegistroSQL("Ticker" & Origen)
   
   DoEvents
   
   VarLinMatrizAcciones = VarLinMatrizAcciones + 1
      
   MatrizAcciones(VarLinMatrizAcciones, 0) = RegistroSQL("Id_Mercados")
   MatrizAcciones(VarLinMatrizAcciones, 1) = RegistroSQL("Ticker" & Origen)
   
   If CDate(RegistroSQL("FechaDesde")) > CDate(VarFechaDesde) Then
   
      MatrizAcciones(VarLinMatrizAcciones, 2) = RegistroSQL("FechaDesde")
      
   Else
      
      MatrizAcciones(VarLinMatrizAcciones, 2) = VarFechaDesde
      
   End If
   
   MatrizAcciones(VarLinMatrizAcciones, 3) = VarFechaHasta

   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub


Public Sub DescargarMercadosImportacionCotizaciones(Origen As String)

TotalProcedimiento = 15
UnidadesProcedimiento = 0

UnidadesTotalesProcedimiento = VarLinMatrizAcciones

UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento

' Recorremos las líneas de MatrizMercados
For i = 1 To VarLinMatrizAcciones
    
    UnidadesProcedimiento = UnidadesProcedimiento + 1

    FrmFiltroImpCotMercados.BarraProgreso.Value = FrmFiltroImpCotMercados.BarraProgreso.Value + UnitarioProcedimiento
    FrmFiltroImpCotMercados.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Descargando CSV cotizaciones " & MatrizAcciones(i, 1)
   
    DoEvents
    
    If Origen = "Yahoo" Then
    
       VarLinkYahoo = "http://ichart.yahoo.com/table.csv?s=%5E[TickerYahoo]&a=[MesDesde]&b=[DiaDesde]&c=[AñoDesde]&d=[MesHasta]&e=[DiaHasta]&f=[AñoHasta]&g=d&ignore=.csv"
       
       VarLinkYahoo = Replace(VarLinkYahoo, "[TickerYahoo]", CStr(MatrizAcciones(i, 1)))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesHasta]", CStr(Month(MatrizAcciones(i, 3))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaHasta]", CStr(Day(MatrizAcciones(i, 3))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoHasta]", CStr(Year(MatrizAcciones(i, 3))))

       VarLinkYahoo = Replace(VarLinkYahoo, "[MesDesde]", CStr(Month(MatrizAcciones(i, 2))) - 1)
       VarLinkYahoo = Replace(VarLinkYahoo, "[DiaDesde]", CStr(Day(MatrizAcciones(i, 2))))
       VarLinkYahoo = Replace(VarLinkYahoo, "[AñoDesde]", CStr(Year(MatrizAcciones(i, 2))))

   
       VarNombreURL = VarLinkYahoo
       VarNombreFichero = CurDir & "\Temp\" & MatrizAcciones(i, 1) & ".csv"
    
    End If
    
    DownloadFile VarNombreURL, VarNombreFichero

       
Next

End Sub

Sub ImportarCotizacionesMercados(IdMercados As String, Fecha As String, Apertura As String, Cierre As String, Maximo As String, Minimo As String, Volumen As String, Origen As String)

' Por si estan abiertas
Set RegistroSQL = Nothing

' Creamos los objetos
Set RegistroSQL = New ADODB.Recordset

' Para arreglar entrada de fecha en formato yyyy-mm-dd o dd/mm/yyyy
If InStr(1, Fecha, "-", vbTextCompare) <> 0 Then

   Fecha = Replace(Fecha, "-", "")
   
Else

   Fecha = Year(Fecha) & Right("00" & Month(Fecha), 2) & Right("00" & Day(Fecha), 2)
   
End If


' Montamos la rutina inicialmente para insertar
RutinaSQL = "INSERT INTO MercadosCotizaciones (Id_Mercados, Fecha, Apertura, Cierre, Maximo, Minimo, Volumen, Origen, ActFecha) VALUES ('[IdMercados]', '[Fecha]', '[Apertura]', '[Cierre]', '[Maximo]', '[Minimo]', '[Volumen]', '[Origen]', GETDATE())"

' Abrimos el recordset de la consulta
RegistroSQL.Open "SELECT Id_Mercados FROM MercadosCotizaciones WHERE Id_Mercados = '" & CStr(IdMercados) & "' AND Fecha = '" & CStr(Replace(Fecha, "-", "")) & "'", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   
   ' Montamos la rutina para actualizar porque ya existe el registro
   RutinaSQL = "UPDATE MercadosCotizaciones SET Apertura = '[Apertura]', Cierre = '[Cierre]', Maximo = '[Maximo]', Minimo = '[Minimo]', Volumen = '[Volumen]', Origen = 'I', ActFecha = GETDATE() WHERE Id_Mercados = '[IdMercados]' AND Fecha = '[Fecha]'"
   
   RegistroSQL.MoveNext
    
Loop

RutinaSQL = Replace(RutinaSQL, "[IdMercados]", CStr(IdMercados))
RutinaSQL = Replace(RutinaSQL, "[Fecha]", CStr(Replace(Fecha, "-", "")))
RutinaSQL = Replace(RutinaSQL, "[Apertura]", CStr(Apertura))
RutinaSQL = Replace(RutinaSQL, "[Cierre]", CStr(Cierre))
RutinaSQL = Replace(RutinaSQL, "[Maximo]", CStr(Maximo))
RutinaSQL = Replace(RutinaSQL, "[Minimo]", CStr(Minimo))
RutinaSQL = Replace(RutinaSQL, "[Volumen]", CStr(Volumen))
RutinaSQL = Replace(RutinaSQL, "[Origen]", CStr(Origen))

Set ResultadoSQL = ConexionSQL.Execute(RutinaSQL)

End Sub

Public Sub CargarArbolMercados()


Dim Zona As String
Dim Pais As String
Dim TickerMercado As String

Zona = ""
TickerMercado = ""
Pais = ""

Dim MiNodo As Node  'Declaramos la variable tipo Node
frmMercadosSegmentados.VistaArbol.Style = 7    'Hacemos que el estilo sea Líneas, +/-, Imagen y Texto

frmMercadosSegmentados.VistaArbol.ImageList = frmMercadosSegmentados.ImageList1

Set MiNodo = frmMercadosSegmentados.VistaArbol.Nodes.Add(, , "SM", "Mercados", "LogoSM")   'No tienen parámetro RelativoA

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
    
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Mercados.Zona, Mercados.TickerYahoo As TickerMercado, Mercados.Nombre As Mercado FROM Mercados ORDER BY Mercados.Zona, Mercados.Nombre", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   If RegistroSQL("Zona") <> Zona Then
   
      Zona = RegistroSQL("Zona")

      Set MiNodo = frmMercadosSegmentados.VistaArbol.Nodes.Add("SM", 4, Left(RegistroSQL("Zona"), 3), RegistroSQL("Zona"), "Zona")
   
   End If
   
      
   Set MiNodo = frmMercadosSegmentados.VistaArbol.Nodes.Add(Left(RegistroSQL("Zona"), 3), 4, RegistroSQL("TickerMercado"), RegistroSQL("TickerMercado") & " - " & RegistroSQL("Mercado"), "Cerrado")

   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub

Public Sub CargarArbolAcciones()


Dim Zona As String
Dim Pais As String
Dim TickerMercado As String

Zona = ""
TickerMercado = ""
Pais = ""

Dim MiNodo As Node  'Declaramos la variable tipo Node
frmArbolAcciones.VistaArbol.Style = 7    'Hacemos que el estilo sea Líneas, +/-, Imagen y Texto

frmArbolAcciones.VistaArbol.ImageList = frmArbolAcciones.ImageList1

Set MiNodo = frmArbolAcciones.VistaArbol.Nodes.Add(, , "SM", "Acciones", "LogoSM")   'No tienen parámetro RelativoA

' Por si estan abiertas
Set ConexionSQL = Nothing
Set RegistroSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection
Set RegistroSQL = New ADODB.Recordset
    
' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
    
' Abrimos el recordset de la tabla Mercados
RegistroSQL.Open "SELECT Mercados.Zona, Mercados.TickerYahoo As TickerMercado, Mercados.Nombre As Mercado, Acciones.TickerYahoo As TickerAccion, Acciones.Nombre As Accion FROM Mercados INNER JOIN AccionesMercados ON AccionesMercados.Id_Mercados = Mercados.Id_Mercados INNER JOIN Acciones ON Acciones.Id_Acciones = AccionesMercados.Id_Acciones GROUP BY Mercados.Zona, Mercados.TickerYahoo, Mercados.Nombre, Acciones.TickerYahoo, Acciones.Nombre ORDER BY Mercados.Zona, Mercados.Nombre, Acciones.Nombre", ConexionSQL, adOpenDynamic, adLockOptimistic

Do While Not RegistroSQL.EOF

   If RegistroSQL("Zona") <> Zona Then
   
      Zona = RegistroSQL("Zona")

      Set MiNodo = frmArbolAcciones.VistaArbol.Nodes.Add("SM", 4, Left(RegistroSQL("Zona"), 3), RegistroSQL("Zona"), "Zona")
   
   End If
   
   If RegistroSQL("TickerMercado") <> TickerMercado Then
 
      TickerMercado = RegistroSQL("TickerMercado")
      
      Set MiNodo = frmArbolAcciones.VistaArbol.Nodes.Add(Left(RegistroSQL("Zona"), 3), 4, RegistroSQL("TickerMercado"), RegistroSQL("TickerMercado") & " - " & RegistroSQL("Mercado"), "Cerrado", "Abierto")
   
   End If
    
   Set MiNodo = frmArbolAcciones.VistaArbol.Nodes.Add(CStr(RegistroSQL("TickerMercado")), 4, RegistroSQL("TickerMercado") & "." & RegistroSQL("TickerAccion"), RegistroSQL("TickerAccion") & " - " & RegistroSQL("Accion"), "Accion")
 

   RegistroSQL.MoveNext
    
Loop

' Cerramos los objetos abiertos para la conexión
RegistroSQL.Close
ConexionSQL.Close

End Sub

Public Sub RecogerAccionesImportacionCotizacionesFSO()

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
          
' Indicamos el porcentaje que total que le corresponde al procedimiento
TotalProcedimiento = 80

' Inicializamos las unidades del procedimiento, las totales y el coste unitario del procedimiento
UnidadesProcedimiento = 0

If VarLinMatrizAcciones <> 0 Then

   UnidadesTotalesProcedimiento = VarLinMatrizAcciones
   UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento
   
End If

' Recorremos las líneas de MatrizMercados
For i = 1 To VarLinMatrizAcciones

    UnidadesProcedimiento = UnidadesProcedimiento + 1

    FrmFiltroImpCotAcciones.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Recogiendo e Importando cotizaciones " & MatrizAcciones(i, 1)
    
    DoEvents
    
    'Montamos el nombre del fichero del mercado
    VarNombreFichero = CurDir & "\Temp\" & CStr(MatrizAcciones(i, 1)) & ".csv"
              
    'Si existe el fichero
    If Len(Dir(VarNombreFichero)) > 0 Then
    
       'Ponemos a cero la variable de control de lineas del fichero
       VarNumeroLineas = 0
       
       ' Referencia al archivo con GetFile
       Set Obj_File = obj_FSO.GetFile(VarNombreFichero)
      
       ' Lo abre con OpenAsTextStream
       Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
      
       ' recorre todo el contenido del fichero
       Do While Not Obj_TextStream.AtEndOfStream
      
          ' lee la linea
          VarTexto = Obj_TextStream.ReadLine
           
          ' agregamos uno al número de lineas leids
          VarNumeroLineas = VarNumeroLineas + 1
      
       Loop
       
       ' cerramos le fichero
       Obj_TextStream.Close
       
       ' asignamos al tamaño original del fichero el número total de líneas del fichero
       TamañoOriginalFichero = VarNumeroLineas
          
       ' calculamos el valor de cada una de las líneas del fichero dentro del total del procedimiento
       UnitarioDetalleProcedimiento = UnitarioProcedimiento / TamañoOriginalFichero
       
       ' inicializamos el número de líneas
       VarNumeroLineas = 0
       
       ' Referencia al archivo con GetFile
       Set Obj_File = obj_FSO.GetFile(VarNombreFichero)
      
       ' Lo abre con OpenAsTextStream
       Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
      
       ' recorre todo el contenido del fichero
       Do While Not Obj_TextStream.AtEndOfStream
      
          ' lee la linea
          VarTexto = Obj_TextStream.ReadLine
          
          'Mientras que exista un separador de campos ","
          While InStr(1, VarTexto, ",", vbTextCompare) <> 0
                    
                ' si todavia hay algun "," en el texto de la línea
                If InStr(1, VarTexto, ",", vbTextCompare) <> 0 Then
                
                   ' agregamos uno al número de campo
                   NumCampo = NumCampo + 1
               
                   ' si el número de campo es mayor que 7
                   If NumCampo > 7 Then NumCampo = 1
                    
                   ' recogemos el valor del campo
                   VarCampo = Left(VarTexto, InStr(1, VarTexto, ",", vbTextCompare) - 1)
                      
                   ' quitamos de vartexto el campo y el separador ","
                   VarTexto = Right(VarTexto, (Len(VarTexto) - InStr(1, VarTexto, ",", vbTextCompare)))
                      
                   ' asignamos le valor del campo al campo que corresponda
                   AsignarValorCampoCotizacion NumCampo, VarCampo
               
                   ' si no existe ningún separador de campo
                   If InStr(1, VarTexto, ",", vbTextCompare) = 0 Then
                   
                      ' agregamos uno al número de campo
                      NumCampo = NumCampo + 1
               
                      ' si el número de campo es mayor que 7
                      If NumCampo > 7 Then NumCampo = 1
               
                      ' cogemos el valor del campo que es lo que queda de la línea
                      VarCampo = VarTexto
                      
                      ' ponemos en blanco el texto de la línea
                      VarTexto = ""
                          
                      ' asignamos le valor del campo al campo que corresponda
                      AsignarValorCampoCotizacion NumCampo, VarCampo
                      
                    
                      ' si es numerico el volumen (para evitar la primera línea que es la cabecera en texto)
                      If IsNumeric(CotVolumen) Then
                      
                         ' Pasamos los datos de la cotización a la función encargada de insertar datos en SQL
                         ImportarCotizacionesAcciones CStr(MatrizAcciones(i, 0)), CotFecha, CotApertura, CotCierre, CotMaximo, CotMinimo, CotVolumen, "I"
                      
                      End If

                   End If
               
                End If
          
          Wend
      
          ' agregamos 1 al número de líneas tratadas del fichero
          VarNumeroLineas = VarNumeroLineas + 1
                       
          ' si el valor que vamos a asignar a la barra de progreso esta dentro de lo permitido
          If FrmFiltroImpCotAcciones.BarraProgreso.Value + UnitarioDetalleProcedimiento <= 100 Then
                      
             ' asignamos valor a la barra de progreso
             FrmFiltroImpCotAcciones.BarraProgreso.Value = FrmFiltroImpCotAcciones.BarraProgreso.Value + UnitarioDetalleProcedimiento
                       
          End If
                      
          ' ponemos el valor que estamos tratando y el porcentaje que queda en la label
          FrmFiltroImpCotAcciones.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Recogiendo e Importando cotizaciones " & MatrizAcciones(i, 1) & " (" & FormatNumber((TamañoOriginalFichero - VarNumeroLineas) / (TamañoOriginalFichero / 100), 2, True, False, True) & "%)"
                      
          ' para forzar el refresco del form
          DoEvents
      
       Loop
       
       ' Cerramos el fichero
       Obj_TextStream.Close
        
       ' Borramos el fichero
       Kill (VarNombreFichero)

    ' Si no existe el fichero
    Else
    
       ' Escribimos en el registro la inexistencia del fichero
       R = EscribirRegistro(RutaRegistro, Now & " - No se han descargado las cotizaciones del valor " & CStr(MatrizAcciones(i, 1)))
     
    End If
    
    ' Asignamos valor a la barra de progreso
    FrmFiltroImpCotAcciones.BarraProgreso.Value = (100 - TotalProcedimiento) + (UnitarioProcedimiento * UnidadesProcedimiento)
    
    ' Forzamos el refresco del form de la barra de progreso
    DoEvents

Next

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close
 
End Sub

Public Sub RecogerMercadosImportacionCotizacionesFSO()

Set ConexionSQL = Nothing
Set ResultadoSQL = Nothing

' Creamos los objetos
Set ConexionSQL = New ADODB.Connection

' Abrimos la base de datos
ConexionSQL.Open FicheroUDLSQL
          
' Indicamos el porcentaje que total que le corresponde al procedimiento
TotalProcedimiento = 80

' Inicializamos las unidades del procedimiento, las totales y el coste unitario del procedimiento
UnidadesProcedimiento = 0
UnidadesTotalesProcedimiento = VarLinMatrizAcciones
UnitarioProcedimiento = TotalProcedimiento / UnidadesTotalesProcedimiento

' Recorremos las líneas de MatrizMercados
For i = 1 To VarLinMatrizAcciones

    UnidadesProcedimiento = UnidadesProcedimiento + 1

    FrmFiltroImpCotMercados.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Recogiendo e Importando cotizaciones " & MatrizAcciones(i, 1)
    
    DoEvents
    
    'Montamos el nombre del fichero del mercado
    VarNombreFichero = CurDir & "\Temp\" & CStr(MatrizAcciones(i, 1)) & ".csv"
              
    'Si existe el fichero
    If Len(Dir(VarNombreFichero)) > 0 Then
    
       'Ponemos a cero la variable de control de lineas del fichero
       VarNumeroLineas = 0
       
       ' Referencia al archivo con GetFile
       Set Obj_File = obj_FSO.GetFile(VarNombreFichero)
      
       ' Lo abre con OpenAsTextStream
       Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
      
       ' recorre todo el contenido del fichero
       Do While Not Obj_TextStream.AtEndOfStream
      
          ' lee la linea
          VarTexto = Obj_TextStream.ReadLine
           
          ' agregamos uno al número de lineas leids
          VarNumeroLineas = VarNumeroLineas + 1
      
       Loop
       
       ' cerramos le fichero
       Obj_TextStream.Close
       
       ' asignamos al tamaño original del fichero el número total de líneas del fichero
       TamañoOriginalFichero = VarNumeroLineas
          
       ' calculamos el valor de cada una de las líneas del fichero dentro del total del procedimiento
       UnitarioDetalleProcedimiento = UnitarioProcedimiento / TamañoOriginalFichero
       
       ' inicializamos el número de líneas
       VarNumeroLineas = 0
       
       ' Referencia al archivo con GetFile
       Set Obj_File = obj_FSO.GetFile(VarNombreFichero)
      
       ' Lo abre con OpenAsTextStream
       Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
      
       ' recorre todo el contenido del fichero
       Do While Not Obj_TextStream.AtEndOfStream
      
          ' lee la linea
          VarTexto = Obj_TextStream.ReadLine
          
          'Mientras que exista un separador de campos ","
          While InStr(1, VarTexto, ",", vbTextCompare) <> 0
                    
                ' si todavia hay algun "," en el texto de la línea
                If InStr(1, VarTexto, ",", vbTextCompare) <> 0 Then
                
                   ' agregamos uno al número de campo
                   NumCampo = NumCampo + 1
               
                   ' si el número de campo es mayor que 7
                   If NumCampo > 7 Then NumCampo = 1
                    
                   ' recogemos el valor del campo
                   VarCampo = Left(VarTexto, InStr(1, VarTexto, ",", vbTextCompare) - 1)
                      
                   ' quitamos de vartexto el campo y el separador ","
                   VarTexto = Right(VarTexto, (Len(VarTexto) - InStr(1, VarTexto, ",", vbTextCompare)))
                      
                   ' asignamos le valor del campo al campo que corresponda
                   AsignarValorCampoCotizacion NumCampo, VarCampo
               
                   ' si no existe ningún separador de campo
                   If InStr(1, VarTexto, ",", vbTextCompare) = 0 Then
                   
                      ' agregamos uno al número de campo
                      NumCampo = NumCampo + 1
               
                      ' si el número de campo es mayor que 7
                      If NumCampo > 7 Then NumCampo = 1
               
                      ' cogemos el valor del campo que es lo que queda de la línea
                      VarCampo = VarTexto
                      
                      ' ponemos en blanco el texto de la línea
                      VarTexto = ""
                          
                      ' asignamos le valor del campo al campo que corresponda
                      AsignarValorCampoCotizacion NumCampo, VarCampo
                      
                      
                      ' si es numerico el volumen (para evitar la primera línea que es la cabecera en texto)
                      If IsNumeric(CotVolumen) Then
                      
                         ' pasamos los datos de la cotización al procedimiento de SQL para que inserte el día en cuestion
                         ImportarCotizacionesMercados CStr(MatrizAcciones(i, 0)), CotFecha, CotApertura, CotCierre, CotMaximo, CotMinimo, CotVolumen, "I"
                      
                      End If

                   End If
               
                End If
          
          Wend
      
          ' agregamos 1 al número de líneas tratadas del fichero
          VarNumeroLineas = VarNumeroLineas + 1
                       
          ' si el valor que vamos a asignar a la barra de progreso esta dentro de lo permitido
          If FrmFiltroImpCotMercados.BarraProgreso.Value + UnitarioDetalleProcedimiento <= 100 Then
                      
             ' asignamos valor a la barra de progreso
             FrmFiltroImpCotMercados.BarraProgreso.Value = FrmFiltroImpCotMercados.BarraProgreso.Value + UnitarioDetalleProcedimiento
                       
          End If
                      
          ' ponemos el valor que estamos tratando y el porcentaje que queda en la label
          FrmFiltroImpCotMercados.LComentario.Caption = UnidadesProcedimiento & " de " & UnidadesTotalesProcedimiento & " - Recogiendo e Importando cotizaciones " & MatrizAcciones(i, 1) & " (" & FormatNumber((TamañoOriginalFichero - VarNumeroLineas) / (TamañoOriginalFichero / 100), 2, True, False, True) & "%)"
                      
          ' para forzar el refresco del form
          DoEvents
      
       Loop
       
       ' Cerramos el fichero
       Obj_TextStream.Close
        
       ' Borramos el fichero
       Kill (VarNombreFichero)
     
    ' Si no existe el fichero
    Else
    
       ' Escribimos en el registro la inexistencia del fichero
       R = EscribirRegistro(RutaRegistro, Now & " - No se han descargado las cotizaciones del mercado " & CStr(MatrizAcciones(i, 1)))
    
    End If
    
    ' Asignamos valor a la barra de progreso
    FrmFiltroImpCotMercados.BarraProgreso.Value = (100 - TotalProcedimiento) + (UnitarioProcedimiento * UnidadesProcedimiento)
    
    ' Forzamos el refresco del form de la barra de progreso
    DoEvents

Next

' Cerramos los objetos abiertos para la conexión
ConexionSQL.Close
 
End Sub
