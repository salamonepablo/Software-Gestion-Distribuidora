

   MEMBER('FEAFIP.clw')                                    ! This is a MEMBER module


   INCLUDE('ABTOOLBA.INC'),ONCE
   INCLUDE('ABWINDOW.INC'),ONCE

                     MAP
                       INCLUDE('FEAFI001.INC'),ONCE        !Local module procedure declarations
                     END


Main PROCEDURE                                             ! Generated from procedure template - Frame

LocalRequest         LONG                                  !
CAE                  STRING(20)                            !
tipocomp             ULONG                                 !
ptovta               ULONG                                 !
Contribuyente        PSTRING(20)                           !
fechacmp             STRING(20)                            !
nro                  ULONG                                 !
pos                  ULONG                                 !
tipo                 ULONG                                 !
URLWSAA              STRING(255)                           !
URLWSW               STRING(255)                           !
CurrentTab           STRING(80)                            !
cResultado           BYTE                                  !
FilesOpened          BYTE                                  !
DisplayDayString STRING('Sunday   Monday   Tuesday  WednesdayThursday Friday   Saturday ')
DisplayDayText   STRING(9),DIM(7),OVER(DisplayDayString)
AppFrame             APPLICATION('Process Customer Orders'),AT(,,400,253),FONT('MS Sans Serif',8,COLOR:Black,),STATUS(-1,80,120,45),SYSTEM,MAX,RESIZE,IMM
                       TOOLBAR,AT(0,0,400,127)
                         OLE,AT(214,24,34,13),USE(?OLE),COMPATIBILITY(020H)
                         END
                         BUTTON('Consulta CUIT'),AT(12,33,85,20),USE(?Button2)
                         BUTTON('Consulta Comprobantes'),AT(11,62,85,20),USE(?Button3)
                         BUTTON('Obtener CAE'),AT(11,7,85,19),USE(?Button1)
                       END
                     END

ThisWindow           CLASS(WindowManager)
Ask                    PROCEDURE(),DERIVED                 ! Method added to host embed code
Init                   PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
Kill                   PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
TakeAccepted           PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
TakeWindowEvent        PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
                     END

Toolbar              ToolbarClass

  CODE
  GlobalResponse = ThisWindow.Run()                        ! Opens the window and starts an Accept Loop

!---------------------------------------------------------------------------
DefineListboxStyle ROUTINE
!|
!| This routine create all the styles to be shared in this window
!| It's called after the window open
!|
!---------------------------------------------------------------------------

ThisWindow.Ask PROCEDURE

  CODE
  IF NOT INRANGE(AppFrame{Prop:Timer},1,100)
    AppFrame{Prop:Timer} = 100
  END
    AppFrame{Prop:StatusText,3} = CLIP(DisplayDayText[(TODAY()%7)+1]) & ', ' & FORMAT(TODAY(),@D4)
    AppFrame{Prop:StatusText,4} = FORMAT(CLOCK(),@T3)
  PARENT.Ask


ThisWindow.Init PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  GlobalErrors.SetProcedureName('Main')
  SELF.Request = GlobalRequest                             ! Store the incoming request
  ReturnValue = PARENT.Init()
  IF ReturnValue THEN RETURN ReturnValue.
  SELF.FirstField = 1
  SELF.VCRRequest &= VCRRequest
  SELF.Errors &= GlobalErrors                              ! Set this windows ErrorManager to the global ErrorManager
  CLEAR(GlobalRequest)                                     ! Clear GlobalRequest after storing locally
  CLEAR(GlobalResponse)
  SELF.AddItem(Toolbar)
  OPEN(AppFrame)                                           ! Open window
  SELF.Opened=True
  Do DefineListboxStyle
  INIMgr.Fetch('Main',AppFrame)                            ! Restore window settings from non-volatile store
  SELF.SetAlerts()
  RETURN ReturnValue


ThisWindow.Kill PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  ReturnValue = PARENT.Kill()
  IF ReturnValue THEN RETURN ReturnValue.
  IF SELF.Opened
    INIMgr.Update('Main',AppFrame)                         ! Save window data to non-volatile store
  END
  GlobalErrors.SetProcedureName
  RETURN ReturnValue


ThisWindow.TakeAccepted PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receive all EVENT:Accepted's
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeAccepted()
    CASE ACCEPTED()
    OF ?Button2
              !Los nombres de los parametros de las funciones se obtienen en www.bitingenieria.com.ar/webhelp
      
              ?OLE{PROP:Create} = 'FEAFIPLib.wsPadron'
              ?OLE{'CUIT'} = 20939802593
              ?OLE{'ModoProduccion'} = false
              If ?OLE{'login("certificado.crt", "clave.key")'} <> 0
                 ?OLE{'CUIT'} = 20939802593
                 BIND('Contribuyente', Contribuyente)
                 if ?OLE{'sfConsultar(30610171601)'} <> 0
                      MESSAGE(?OLE{'Contribuyente.nombre'})
                 else
                      MESSAGE(?OLE{'ErrorDesc'})
                 end
              Else
                  MESSAGE(?OLE{'ErrorDesc'})
              End
    OF ?Button3
              !Los nombres de los parametros de las funciones se obtienen en www.bitingenieria.com.ar/webhelp
              
      
                ! URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
              URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms'
                ! Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
              URLWSW = 'https://wswhomo.afip.gov.ar/wsfev1/service.asmx'
                ! Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
      
              Ptovta = 3
              Tipocomp = 1 !Factura A
              nro = 8500
      
              ?OLE{PROP:Create} = 'FEAFIPLib.wsfev1'
              ?OLE{'CUIT'} = 20939802593
              ?OLE{'URL'} = URLWSW
      
              If ?OLE{'login("certificado.crt", "clave.key","' & URLWSAA & '")'} <> 0
                  ?OLE{'SFRecuperaLastCMP(' & Ptovta & ',' & Tipocomp & ')'}
                  nro = ?OLE{'SFLastCMP'}
                  If ?OLE{'SFCmpConsultar(' & Tipocomp & ',' & Ptovta & ',' & nro & ')'} <> 0
                      CAE = ?OLE{'CmpConsultarCbte.CodAutorizacion'}
                      !! Para obtener los valores de los campos a continuación se debe usar la misma sintaxis que con el CAE
                      !!CmpConsultarCbte.Concepto
                      !!CmpConsultarCbte.DocTipo
                      !!CmpConsultarCbte.DocNro
                      !!CmpConsultarCbte.CbteDesde
                      !!CmpConsultarCbte.CbteHasta
                      !!CmpConsultarCbte.CbteFch
                      !!CmpConsultarCbte.ImpTotal
                      !!CmpConsultarCbte.ImpTotConc
                      !!CmpConsultarCbte.ImpNeto
                      !!CmpConsultarCbte.ImpOpEx
                      !!CmpConsultarCbte.ImpTrib
                      !!CmpConsultarCbte.ImpIVA
                      !!CmpConsultarCbte.FchServDesde
                      !!CmpConsultarCbte.FchServHasta
                      !!CmpConsultarCbte.FchVtoPago
                      !!CmpConsultarCbte.MonId
                      !!CmpConsultarCbte.MonCotiz
                      !!CmpConsultarCbte.CbtesAsocCount
                      !!CmpConsultarCbte.TributosCount
                      !!CmpConsultarCbte.IvaCount
                      !!CmpConsultarCbte.OpcionalesCount
                      !!CmpConsultarCbte.Resultado
                      !!CmpConsultarCbte.CodAutorizacion
                      !!CmpConsultarCbte.EmisionTipo
                      !!CmpConsultarCbte.FchVto
                      !!CmpConsultarCbte.FchProceso
                      !!CmpConsultarCbte.ObservacionesCount
                      !!CmpConsultarCbte.PtoVta
                      !!CmpConsultarCbte.CbteTipo
      
                      ! Ver en www.bitingenieria.com.ar/doc/feafip/FEAFIPLib_TLB.IComprobante.html
                      LOOP I# = 0 TO ?OLE{'CmpConsultarCbte.IvaCount'} - 1
                        BaseImp$ = ?OLE{'CmpConsultarCbte.Iva(' & I# & ').BaseImp'}
                        Importe$ = ?OLE{'CmpConsultarCbte.Iva(' & I# & ').Importe'}
                        IdIva$ = ?OLE{'CmpConsultarCbte.Iva(' & I# & ').Id'}
                      END
      
                      MESSAGE(CAE)
                  else
                    MESSAGE(?OLE{'ErrorDesc'})
                  End
              Else
                  MESSAGE(?OLE{'ErrorDesc'})
              End
    OF ?Button1
              !Los nombres de los parametros de las funciones se obtienen en www.bitingenieria.com.ar/webhelp
              
      
                ! URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
              URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms'
                ! Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
              URLWSW = 'https://wswhomo.afip.gov.ar/wsfev1/service.asmx'
                ! Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
      
              Ptovta = 116
              Tipocomp = 1 !Factura A(Ver excel referencias codigos AFIP)
              fechacmp = year(today()) & Format(month(today()), @N02) & Format(day(today()),@N02) ! Tomo la fecha actual como ejemplo
      
              ?OLE{PROP:Create} = 'FEAFIPLib.wsfev1'
              ?OLE{'CUIT'} = 20939802593
              ?OLE{'URL'} = URLWSW
              ?OLE{'Depurar'} = true
      
              If ?OLE{'login("certificado.crt", "clave.key","' & URLWSAA & '")'} <> 0
                  If ?OLE{'SFRecuperaLastCMP(' & Ptovta & ', ' & Tipocomp & ')'} <> 0
                    nro = ?OLE{'SFLastCmp'} !Devolucion el ultimo comprobante
                  else
                    MESSAGE(?OLE{'ErrorDesc'})
                  End
                    nro = nro + 1
                    ?OLE{'Reset()'}
                    ?OLE{'AgregaFactura(1, 80, 30702637895,' & nro & ',' & nro & ',' & clip(fechacmp) &  ', "121,10", 0, 100, 0, "", "", "", "PES", 1)'}
                    ?OLE{'AgregaIVA(5, 100, 21)'}  !Ver excel referencias codigos AFIP
                    If ?OLE{'Autorizar(' & Ptovta & ',' & Tipocomp & ')'} <> 0
                      If ?OLE{'SFresultado(0)'} = 'A'
                        MESSAGE('Felicitaciones Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:' & ?OLE{'SFCAE(0)'} & ' Vencimiento: ' & ?OLE{'SFVencimiento(0)'})
                      Else
                        ! observaciones
                        MESSAGE(?OLE{'AutorizarRespuestaObs(0)'})
                      End
                    Else
                       MESSAGE(?OLE{'ErrorDesc'})
                    End
              Else
                  MESSAGE(?OLE{'ErrorDesc'})
              End
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue


ThisWindow.TakeWindowEvent PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receives all window specific events
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeWindowEvent()
    CASE EVENT()
    OF Event:Timer
      AppFrame{Prop:StatusText,3} = CLIP(DisplayDayText[(TODAY()%7)+1]) & ', ' & FORMAT(TODAY(),@D4)
      AppFrame{Prop:StatusText,4} = FORMAT(CLOCK(),@T3)
    ELSE
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue

