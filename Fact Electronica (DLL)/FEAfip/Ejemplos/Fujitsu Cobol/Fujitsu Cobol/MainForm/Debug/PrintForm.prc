000001 IDENTIFICATION  DIVISION.
000002* PrintForm.
000003 PROGRAM-ID.     PrintForm.
000004 ENVIRONMENT     DIVISION.
000005 CONFIGURATION   SECTION.
000006 POW-REPOSITORY.
000007     CLASS  AMethodSetPrintForm AS "TLB=C:\Users\amiranda\Documents\Embarcadero\Studio\Projects\FEAFIP\Ejemplo\Fujitsu Cobol\Fujitsu Cobol\MainForm\Debug\~build.tlb,{3F3C3B35-0E51-48AA-8BBF-CD2E9C80E5FA},Fujitsu-PcobForm-4"
000008     CLASS  AMixed-DCfGWnd-Main-with-DCfGroupItem-Main AS "TLB=C:\Users\amiranda\Documents\Embarcadero\Studio\Projects\FEAFIP\Ejemplo\Fujitsu Cobol\Fujitsu Cobol\MainForm\Debug\~build.tlb,{EA0440F7-183F-11D1-95C2-00A0C90D6AFE},Fujitsu-PcobFormWnd-4"
000009     CLASS  AMixed-DCmPush-Main-with-DCfGroupItem-Main AS "TLB=C:\Users\amiranda\Documents\Embarcadero\Studio\Projects\FEAFIP\Ejemplo\Fujitsu Cobol\Fujitsu Cobol\MainForm\Debug\~build.tlb,{569A2714-3D90-11D2-B17E-00A0C92DE141},Fujitsu-PcobCommandButton-4"
000010 .
000011 SPECIAL-NAMES.
000012 REPOSITORY.
000013*<SCRIPT DIVISION="ENVIRONMENT", SECTION="CONFIGURATION", PARAGRAPH="REPOSITORY">
000014     CLASS COM AS "*OLE".
000015*</SCRIPT>
000016 .
000017 INPUT-OUTPUT    SECTION.
000018 FILE-CONTROL.
000019 DATA            DIVISION.
000020 BASED-STORAGE   SECTION.
000021 FILE            SECTION.
000022 WORKING-STORAGE SECTION.
000023 CONSTANT        SECTION.
000024 LINKAGE         SECTION.
000025 01  POW-FORM IS GLOBAL.
000026   02  POW-SELF OBJECT REFERENCE AMethodSetPrintForm.
000027   02  POW-SUPER  PIC X(4).
000028   02  POW-THIS OBJECT REFERENCE AMethodSetPrintForm.
000029   02  CmCommand1 OBJECT REFERENCE AMixed-DCmPush-Main-with-DCfGroupItem-Main.
000030 01  PrintForm REDEFINES POW-FORM GLOBAL OBJECT REFERENCE AMethodSetPrintForm.
000031 01  POW-CONTROL-ID PIC S9(9) COMP-5.
000032 01  POW-EVENT-ID   PIC S9(9) COMP-5.
000033 01  POW-OLE-PARAM  PIC X(4).
000034 01  POW-OLE-RETURN PIC X(4).
000035 PROCEDURE       DIVISION USING POW-FORM POW-CONTROL-ID POW-EVENT-ID POW-OLE-PARAM POW-OLE-RETURN.
000036     EVALUATE POW-CONTROL-ID
000037     WHEN 117440517
000038     EVALUATE POW-EVENT-ID
000039     WHEN -600
000040       CALL "POW-SCRIPTLET1"
000041     END-EVALUATE
000042     END-EVALUATE
000043     EXIT PROGRAM.
000044 IDENTIFICATION  DIVISION.
000045* CmCommand1-Click.
000046 PROGRAM-ID.     POW-SCRIPTLET1.
000047*<SCRIPT DIVISION="PROCEDURE", CONTROL="CmCommand1", EVENT="Click", POW-NAME="SCRIPTLET1", TYPE="ETC">
000048 ENVIRONMENT     DIVISION.
000049 DATA            DIVISION.
000050 WORKING-STORAGE SECTION.
000051 01 OBJ-FEAFIP    OBJECT REFERENCE COM.
000052 01 PROGID-FEAFIP PIC X(8192) VALUE "FEAFIPLib.wsfev1".
000053 01 CERTIFICADO PIC X(8192) VALUE "certificado.crt".
000054 01 CLAVE PIC X(8192) VALUE "clave.key".
000055 01 URLWSAA PIC X(8192) VALUE "https://wsaahomo.afip.gov.ar/ws/services/LoginCms".
000056 01 URLWSW PIC X(8192) VALUE "https://wswhomo.afip.gov.ar/wsfev1/service.asmx".
000057 01 IS-OK PIC S9(4) COMP-5.
000058 01 ERROR-DESC PIC X(8192).
000059 01 PtoVta PIC S9(9) COMP-5 VALUE 100.
000060 01 TipoComp PIC S9(9) COMP-5 VALUE 1.
000061 01 CUIT COMP-2 VALUE 20939802593.
000062 01 NRO COMP-2.
000063 01 Concepto PIC S9(9) COMP-5 VALUE 1.
000064 01 DocTipo PIC S9(9) COMP-5 VALUE 80.
000065 01 DocNro COMP-2 VALUE 30702637895.
000066 01 CbteFch PIC X(8192) VALUE "20161007".
000067 01 Imptotal COMP-2 VALUE 121.
000068 01 ImpTotalConc COMP-2 VALUE 0.
000069 01 ImpNeto COMP-2 VALUE 100.
000070 01 ImpOpEx COMP-2 VALUE 0.
000071 01 FechaServDesde PIC X(8192) VALUE " ".
000072 01 FechaServHasta PIC X(8192) VALUE " ".
000073 01 FechaVencPago PIC X(8192) VALUE " ".
000074 01 MonId PIC X(8192) VALUE "PES".
000075 01 MonCotiz COMP-2 VALUE 1.
000076 01 Id-imp PIC S9(9) COMP-5 VALUE 5.
000077 01 BaseImp COMP-2 VALUE 100.
000078 01 Importe COMP-2 VALUE 21.
000079 01 CAE PIC X(8192).
000080 01 Vencimiento PIC X(8192).
000081 01 invoice-index PIC S9(9) COMP-5 VALUE 0.
000082 01 Resultado PIC X(8192).
000083 PROCEDURE       DIVISION.
000084     invoke COM "CREATE-OBJECT" using PROGID-FEAFIP
000085                                returning OBJ-FEAFIP.
000086     invoke OBJ-FEAFIP "SET-CUIT" using CUIT.
000087     invoke OBJ-FEAFIP "SET-URL" using URLWSW
000088     invoke OBJ-FEAFIP "Login" using CERTIFICADO
000089                                     CLAVE
000090                                     URLWSAA.
000091     invoke OBJ-FEAFIP "GET-ErrorCode" returning IS-OK.
000092     IF IS-OK = 0 THEN
000093        invoke OBJ-FEAFIP "SFRecuperaLastCmp" using PtoVta
000094                                                    TipoComp
000095        invoke OBJ-FEAFIP "GET-ErrorCode" returning IS-OK
000096        IF IS-OK = 0 THEN                                    
000097           invoke OBJ-FEAFIP "GET-SFLastCMP" returning NRO
000098           ADD 1 TO NRO
000099           invoke OBJ-FEAFIP "AgregaFactura" using Concepto 
000100                                                   DocTipo
000101                                                   DocNro
000102                                                   NRO
000103                                                   NRO
000104                                                   CbteFch
000105                                                   Imptotal
000106                                                   ImpTotalConc
000107                                                   ImpNeto
000108                                                   ImpOpEx
000109                                                   FechaServDesde
000110                                                   FechaServHasta
000111                                                   FechaVencPago
000112                                                   MonId
000113                                                   MonCotiz
000114           invoke OBJ-FEAFIP "AgregaIVA" using Id-imp
000115                                                   BaseImp
000116                                                   Importe
000117           invoke OBJ-FEAFIP "Autorizar" using PtoVta
000118                                               TipoComp 
000119           invoke OBJ-FEAFIP "GET-ErrorCode" returning IS-OK
000120           IF IS-OK = 0 THEN                         
000121              INVOKE  OBJ-FEAFIP "GET-SFResultado" using invoice-index RETURNING Resultado
000122              IF Resultado = "A" THEN
000123*                Aqui es donde se obtiene el CAE 
000124                 INVOKE OBJ-FEAFIP "GET-SFCAE" using invoice-index returning CAE
000125                 INVOKE OBJ-FEAFIP "GET-SFVencimiento" using invoice-index returning Vencimiento
000126                 INVOKE pow-self "Displaymessage" using CAE
000127              ELSE   
000128                 INVOKE OBJ-FEAFIP "AutorizarRespuestaObs" using invoice-index returning ERROR-DESC
000129                 INVOKE pow-self "Displaymessage" using ERROR-DESC
000130              END-IF   
000131           ELSE
000132              INVOKE pow-self "Displaymessage" using ERROR-DESC
000133           END-IF
000134        ELSE
000135           INVOKE  OBJ-FEAFIP "GET-ErrorDesc" RETURNING ERROR-DESC
000136           INVOKE pow-self "Displaymessage" using ERROR-DESC
000137        END-IF
000138     ELSE
000139        INVOKE  OBJ-FEAFIP "GET-ErrorDesc" RETURNING ERROR-DESC
000140        INVOKE pow-self "Displaymessage" using ERROR-DESC
000141     END-IF.
000142*</SCRIPT>
000143 END PROGRAM     POW-SCRIPTLET1.
000144 END PROGRAM     PrintForm.
