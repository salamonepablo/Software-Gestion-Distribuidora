unit FEAFIPLib_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 04/02/2021 17:39:36 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Users\amiranda\Documents\Embarcadero\Studio\Projects\FEAFIP\Win32\Debug\feafip.dll (1)
// LIBID: {DAE2CF79-E2C1-40C0-90CD-43C9688F108E}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  FEAFIPLibMajorVersion = 1;
  FEAFIPLibMinorVersion = 1;

  LIBID_FEAFIPLib: TGUID = '{DAE2CF79-E2C1-40C0-90CD-43C9688F108E}';

  IID_Iwsaa: TGUID = '{47BE3547-1C9B-4BCA-9F4E-A65234F2C129}';
  CLASS_wsaa: TGUID = '{33B45EE7-0219-4BE1-A9CA-2B57CA4FD209}';
  IID_Iwsfexv1: TGUID = '{10891378-BAE5-4F40-AF39-70C54F4E8175}';
  CLASS_wsfexv1: TGUID = '{CBC36AD9-1D16-4590-A82C-2ED017AAAB4C}';
  IID_Iwsfev1: TGUID = '{E0A95BBC-E328-4AA6-84E2-405C10AD41A2}';
  CLASS_wsfev1: TGUID = '{6804CFD5-32DD-43AE-A463-CB64FCBE32D2}';
  IID_Iwsbfev1: TGUID = '{A5C9683D-3D72-4392-AD49-A4DFB83D8C63}';
  CLASS_wsbfev1: TGUID = '{2E472E22-AD8A-4071-8C62-D2D9B8CE47D3}';
  IID_Iwsmtxca: TGUID = '{C297BD2B-A528-446B-BF55-FAF195383E0E}';
  CLASS_wsmtxca: TGUID = '{C3DD12A3-EAA2-4F45-8F5D-4A25CBD19838}';
  IID_Iwsseg: TGUID = '{B1E85685-67E8-4B99-B8B6-85A6138E4DD0}';
  CLASS_wsseg: TGUID = '{5B4092EF-B311-4CDD-A9F8-61A0AEC7E54C}';
  IID_IwsPadron: TGUID = '{0CEB0878-6393-4701-8C86-2CA793CDCB0D}';
  CLASS_wsPadron: TGUID = '{F57D2D12-E231-4AF7-BB54-3CDDFB52713B}';
  IID_IComprobante: TGUID = '{DC4152DF-68E8-4C5C-804F-22B28CF4C726}';
  CLASS_Comprobante: TGUID = '{A9B8A44F-99B4-4D18-8A54-A66CC5C39BEB}';
  IID_ICbteAsoc: TGUID = '{43E44C59-376E-4A27-93D2-ADC712D2BA2E}';
  IID_ITributo: TGUID = '{8C1BE2D0-B8B0-442E-A8D6-8BBBE941DB0C}';
  IID_IAlicIva: TGUID = '{ADE1B3EE-2618-461B-B8D3-F048B400330A}';
  IID_IOpcional: TGUID = '{7689C644-3F89-44FE-97CF-EAF233A262C8}';
  CLASS_CbteAsoc: TGUID = '{259527D7-6AE5-411F-89EC-9C9A480A41F9}';
  CLASS_Tributo: TGUID = '{FD8F306C-CE28-460C-810C-57CE15C35A37}';
  CLASS_AlicIva: TGUID = '{DA06EFB8-3B19-4061-A544-157036E2CB57}';
  CLASS_Opcional: TGUID = '{74473C21-FABC-49EB-B268-7D6B33D8C728}';
  IID_IObs: TGUID = '{3417F5A9-B0F6-4CF9-B30B-055E17860895}';
  CLASS_Obs: TGUID = '{C778E764-6411-4086-A488-14FC22B1BA4A}';
  IID_IContribuyente: TGUID = '{19A25CC6-4F15-4C2E-AF88-7AD7901B23A9}';
  CLASS_Contribuyente: TGUID = '{0214EC04-2B59-4CDA-BE4F-6212C9B65F02}';
  IID_IDomicilio: TGUID = '{EC378410-896F-4CF2-84A8-53E61AE3D6CF}';
  CLASS_Domicilio: TGUID = '{0370A743-18E0-424E-8124-9CA27A80EB16}';
  IID_IwsPadronARBA: TGUID = '{924DCE98-B918-42E4-A00A-76FD1D8D483A}';
  CLASS_wsPadronARBA: TGUID = '{6F042FCF-8D78-498B-8630-61346537279F}';
  IID_IConsultaAlicuotaRespuesta: TGUID = '{2589E4FF-0788-4FEF-9565-0F05095F1356}';
  CLASS_ConsultaAlicuotaRespuesta: TGUID = '{E17927E3-019B-4B2E-BB5B-CD6DA8A61F59}';
  IID_ICertificado: TGUID = '{CAD1F637-CD57-45DF-8A39-EB2227E34D93}';
  CLASS_Certificado: TGUID = '{189DA0FB-8B57-4C51-834E-666BE83E5878}';
  IID_Iwscdc: TGUID = '{201C6546-D660-4171-A3D3-839583F7969E}';
  CLASS_wscdc: TGUID = '{FDFE6850-8AE7-4B58-8D03-7655A9F28402}';
  IID_IBarcode: TGUID = '{01F6CFB9-A47D-401E-8A89-1C3962BB9364}';
  CLASS_Barcode: TGUID = '{C53215BA-E553-4742-8D17-193B041996F9}';
  IID_Iwsct: TGUID = '{161A74B4-F8B8-408F-934B-2D2D32E492E2}';
  CLASS_wsct: TGUID = '{0C0A8678-5679-4F36-A995-85DA19D90CF5}';
  IID_Iwsfecred: TGUID = '{32EF8E70-4CB3-40FD-A66C-BBB03E147C37}';
  CLASS_wsfecred: TGUID = '{A9EFDBFE-92AC-4D25-B52E-053810AAEB03}';
  IID_IIdCtaCteTy: TGUID = '{C9194512-99E1-4404-85AB-6218E498CEED}';
  CLASS_IdCtaCteTy: TGUID = '{B27AFDD3-1232-49BB-B143-8D6B9CDD2AC5}';
  IID_IIdComprobanteTy: TGUID = '{3ABD3582-6764-4A05-BFDE-CFED3D4A1143}';
  CLASS_IdComprobanteTy: TGUID = '{02CD78F9-A667-49EA-8508-CFC2F476AB45}';
  IID_IAceptarFECredRequestTy: TGUID = '{F0324362-5DE0-4A53-B253-D18C37D5FD5C}';
  CLASS_AceptarFECredRequestTy: TGUID = '{4516C1D1-6EC6-4017-BCB7-75E2C45C45D3}';
  IID_IConsultarCmpReturnTy: TGUID = '{BB315EBC-4D6F-4542-BC1D-B4878E91A9EF}';
  CLASS_ConsultarCmpReturnTy: TGUID = '{C2B0F281-F728-47D0-ACD5-30F614857DD6}';
  IID_IComprobanteTy: TGUID = '{E4EF8C5E-D0D5-4C0A-A1AA-AFDC93A8D4E7}';
  CLASS_ComprobanteTy: TGUID = '{5088161C-D4C5-4804-B3D2-ADBB7105F6CD}';
  IID_ISubtotalIVATy: TGUID = '{873012B5-0A40-440E-9F18-ED81C3C7AD4F}';
  CLASS_SubtotalIVATy: TGUID = '{6AB7D1ED-1C7B-498D-B0CB-639178F5E98F}';
  IID_IOtroTributoTy: TGUID = '{96443A17-3274-4493-A940-92F4FE8F4D98}';
  CLASS_OtroTributoTy: TGUID = '{DAA88209-E23C-4E71-922B-1F4A2789119A}';
  IID_IItemTy: TGUID = '{572B401B-91D9-46CA-85A7-ED286B14693B}';
  CLASS_ItemTy: TGUID = '{2FA6E1A5-A1F8-4A90-9B4F-2C8A722AD481}';
  IID_IMotivoRechazoTy: TGUID = '{21AC85E6-B7A9-487F-BCBC-19E18AE05D42}';
  CLASS_MotivoRechazoType: TGUID = '{519CBF25-50EE-42E6-A732-A7CC2AC99A1C}';
  IID_IInformarFacturaAgtDptoCltvRequestTy: TGUID = '{3DEF5AFE-1202-4B92-BEBB-4B006E29C02F}';
  CLASS_InformarFacturaAgtDptoCltvRequestTy: TGUID = '{CD68CB93-ECDF-42BE-9854-02FA02E8B30C}';
  IID_IRechazarFECredRequestTy: TGUID = '{30EBD9FB-D607-484D-A5E8-8AD7522DA407}';
  CLASS_RechazarFECredRequestTy: TGUID = '{28C2A13F-C5BE-49BE-8DA4-635CF9529B93}';
  IID_IconsultarObligadoRecepcionReturnTy: TGUID = '{2C7111F1-8465-43EB-9110-9303B3961AC3}';
  CLASS_consultarObligadoRecepcionReturnTy: TGUID = '{6A3EF1B1-656A-43F8-B80A-E2C375D42615}';
  IID_IConsultarMontoObligadoRecepcionReturnTy: TGUID = '{24CDB620-0B79-4E7D-943A-3F55F1E26C95}';
  CLASS_ConsultarMontoObligadoRecepcionReturnTy: TGUID = '{0DB7F725-8B5C-4024-8B15-A5CD5588B8A6}';
  IID_IConsultarCtasCtesReturnTy: TGUID = '{9E84530B-FB93-4225-BB57-8BA22738ED6A}';
  CLASS_ConsultarCtasCtesReturnTy: TGUID = '{87058F67-F2AA-4258-A232-B7B4DB2D5AD4}';
  IID_IInfoCtaCteTy: TGUID = '{AF2B653F-D0F7-40CA-BAF7-B8A30A2F03E0}';
  CLASS_InfoCtaCteTy: TGUID = '{B1311E00-BDEE-41B0-A6A5-C7D51F70A30E}';
  IID_IConsultarCtaCteReturnTy: TGUID = '{A70570EB-65D9-4117-A7AB-A57B902E3407}';
  CLASS_ConsultarCtaCteReturnTy: TGUID = '{F4BAAC5D-D8F1-4435-B8DC-085AF2DF2803}';
  IID_IFEGenerador: TGUID = '{429B1793-8E1A-4170-AFA9-7E499F3F6076}';
  CLASS_FEGenerador: TGUID = '{1AA1A982-6EA9-4666-89DF-AE40261EE44C}';
  IID_IQr: TGUID = '{75D6A95C-92FF-4A01-AD58-0AF819349713}';
  CLASS_Qr: TGUID = '{0F88D209-5133-4EF5-92D8-9DADB04695C8}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TipoComprobante
type
  TipoComprobante = TOleEnum;
const
  tcFacturaA = $00000001;
  tcNotaDebitoA = $00000002;
  tcNotaCreditoA = $00000003;
  tcFacturaB = $00000006;
  tcNotaDebitoB = $00000007;
  tcNotaCreditoB = $00000008;
  tcFacturaC = $0000000B;
  tcNotaDebitoC = $0000000C;
  tcNotaCreditoC = $0000000D;

// Constants for enum UnidadesDeMedida
type
  UnidadesDeMedida = TOleEnum;
const
  const0 = $00000000;
  Const1 = $00000001;

// Constants for enum TipoResponsable
type
  TipoResponsable = TOleEnum;
const
  trInscripto = $00000001;
  trNoInscripto = $00000002;
  trNoResponsable = $00000003;
  trExento = $00000004;
  trConsumidorFinal = $00000005;
  trMonotributo = $00000006;
  trNoCategorizado = $00000007;
  trProveedorExterior = $00000008;
  trClienteExterior = $00000009;
  trIVALiberado = $0000000A;
  trInscriptoAgentePerc = $0000000B;
  trPequenioEventual = $0000000C;
  trMonotribSocial = $0000000D;
  trPequenioContribSocial = $0000000E;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  Iwsaa = interface;
  IwsaaDisp = dispinterface;
  Iwsfexv1 = interface;
  Iwsfexv1Disp = dispinterface;
  Iwsfev1 = interface;
  Iwsfev1Disp = dispinterface;
  Iwsbfev1 = interface;
  Iwsbfev1Disp = dispinterface;
  Iwsmtxca = interface;
  IwsmtxcaDisp = dispinterface;
  Iwsseg = interface;
  IwssegDisp = dispinterface;
  IwsPadron = interface;
  IwsPadronDisp = dispinterface;
  IComprobante = interface;
  IComprobanteDisp = dispinterface;
  ICbteAsoc = interface;
  ICbteAsocDisp = dispinterface;
  ITributo = interface;
  ITributoDisp = dispinterface;
  IAlicIva = interface;
  IAlicIvaDisp = dispinterface;
  IOpcional = interface;
  IOpcionalDisp = dispinterface;
  IObs = interface;
  IObsDisp = dispinterface;
  IContribuyente = interface;
  IContribuyenteDisp = dispinterface;
  IDomicilio = interface;
  IDomicilioDisp = dispinterface;
  IwsPadronARBA = interface;
  IwsPadronARBADisp = dispinterface;
  IConsultaAlicuotaRespuesta = interface;
  IConsultaAlicuotaRespuestaDisp = dispinterface;
  ICertificado = interface;
  ICertificadoDisp = dispinterface;
  Iwscdc = interface;
  IwscdcDisp = dispinterface;
  IBarcode = interface;
  IBarcodeDisp = dispinterface;
  Iwsct = interface;
  IwsctDisp = dispinterface;
  Iwsfecred = interface;
  IwsfecredDisp = dispinterface;
  IIdCtaCteTy = interface;
  IIdCtaCteTyDisp = dispinterface;
  IIdComprobanteTy = interface;
  IIdComprobanteTyDisp = dispinterface;
  IAceptarFECredRequestTy = interface;
  IAceptarFECredRequestTyDisp = dispinterface;
  IConsultarCmpReturnTy = interface;
  IConsultarCmpReturnTyDisp = dispinterface;
  IComprobanteTy = interface;
  IComprobanteTyDisp = dispinterface;
  ISubtotalIVATy = interface;
  ISubtotalIVATyDisp = dispinterface;
  IOtroTributoTy = interface;
  IOtroTributoTyDisp = dispinterface;
  IItemTy = interface;
  IItemTyDisp = dispinterface;
  IMotivoRechazoTy = interface;
  IMotivoRechazoTyDisp = dispinterface;
  IInformarFacturaAgtDptoCltvRequestTy = interface;
  IInformarFacturaAgtDptoCltvRequestTyDisp = dispinterface;
  IRechazarFECredRequestTy = interface;
  IRechazarFECredRequestTyDisp = dispinterface;
  IconsultarObligadoRecepcionReturnTy = interface;
  IconsultarObligadoRecepcionReturnTyDisp = dispinterface;
  IConsultarMontoObligadoRecepcionReturnTy = interface;
  IConsultarMontoObligadoRecepcionReturnTyDisp = dispinterface;
  IConsultarCtasCtesReturnTy = interface;
  IConsultarCtasCtesReturnTyDisp = dispinterface;
  IInfoCtaCteTy = interface;
  IInfoCtaCteTyDisp = dispinterface;
  IConsultarCtaCteReturnTy = interface;
  IConsultarCtaCteReturnTyDisp = dispinterface;
  IFEGenerador = interface;
  IFEGeneradorDisp = dispinterface;
  IQr = interface;
  IQrDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  wsaa = Iwsaa;
  wsfexv1 = Iwsfexv1;
  wsfev1 = Iwsfev1;
  wsbfev1 = Iwsbfev1;
  wsmtxca = Iwsmtxca;
  wsseg = Iwsseg;
  wsPadron = IwsPadron;
  Comprobante = IComprobante;
  CbteAsoc = ICbteAsoc;
  Tributo = ITributo;
  AlicIva = IAlicIva;
  Opcional = IOpcional;
  Obs = IObs;
  Contribuyente = IContribuyente;
  Domicilio = IDomicilio;
  wsPadronARBA = IwsPadronARBA;
  ConsultaAlicuotaRespuesta = IConsultaAlicuotaRespuesta;
  Certificado = ICertificado;
  wscdc = Iwscdc;
  Barcode = IBarcode;
  wsct = Iwsct;
  wsfecred = Iwsfecred;
  IdCtaCteTy = IIdCtaCteTy;
  IdComprobanteTy = IIdComprobanteTy;
  AceptarFECredRequestTy = IAceptarFECredRequestTy;
  ConsultarCmpReturnTy = IConsultarCmpReturnTy;
  ComprobanteTy = IComprobanteTy;
  SubtotalIVATy = ISubtotalIVATy;
  OtroTributoTy = IOtroTributoTy;
  ItemTy = IItemTy;
  MotivoRechazoType = IMotivoRechazoTy;
  InformarFacturaAgtDptoCltvRequestTy = IInformarFacturaAgtDptoCltvRequestTy;
  RechazarFECredRequestTy = IRechazarFECredRequestTy;
  consultarObligadoRecepcionReturnTy = IconsultarObligadoRecepcionReturnTy;
  ConsultarMontoObligadoRecepcionReturnTy = IConsultarMontoObligadoRecepcionReturnTy;
  ConsultarCtasCtesReturnTy = IConsultarCtasCtesReturnTy;
  InfoCtaCteTy = IInfoCtaCteTy;
  ConsultarCtaCteReturnTy = IConsultarCtaCteReturnTy;
  FEGenerador = IFEGenerador;
  Qr = IQr;


// *********************************************************************//
// Interface: Iwsaa
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {47BE3547-1C9B-4BCA-9F4E-A65234F2C129}
// *********************************************************************//
  Iwsaa = interface(IDispatch)
    ['{47BE3547-1C9B-4BCA-9F4E-A65234F2C129}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString; const Servicio: WideString): OLE_CANCELBOOL; safecall;
    function Get_Token: WideString; safecall;
    function Get_Sign: WideString; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_CUIT: WideString; safecall;
    procedure Set_CUIT(const Value: WideString); safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function Get_Proxy: WideString; safecall;
    procedure Set_Proxy(const Value: WideString); safecall;
    function Get_ProxyUserName: WideString; safecall;
    procedure Set_ProxyUserName(const Value: WideString); safecall;
    function Get_ProxyPassword: WideString; safecall;
    procedure Set_ProxyPassword(const Value: WideString); safecall;
    function Get_ProxyEnabled: OLE_CANCELBOOL; safecall;
    procedure Set_ProxyEnabled(Value: OLE_CANCELBOOL); safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    property Token: WideString read Get_Token;
    property Sign: WideString read Get_Sign;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property CUIT: WideString read Get_CUIT write Set_CUIT;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property Proxy: WideString read Get_Proxy write Set_Proxy;
    property ProxyUserName: WideString read Get_ProxyUserName write Set_ProxyUserName;
    property ProxyPassword: WideString read Get_ProxyPassword write Set_ProxyPassword;
    property ProxyEnabled: OLE_CANCELBOOL read Get_ProxyEnabled write Set_ProxyEnabled;
  end;

// *********************************************************************//
// DispIntf:  IwsaaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {47BE3547-1C9B-4BCA-9F4E-A65234F2C129}
// *********************************************************************//
  IwsaaDisp = dispinterface
    ['{47BE3547-1C9B-4BCA-9F4E-A65234F2C129}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString; const Servicio: WideString): OLE_CANCELBOOL; dispid 101;
    property Token: WideString readonly dispid 104;
    property Sign: WideString readonly dispid 105;
    property ErrorCode: Integer readonly dispid 102;
    property ErrorDesc: WideString readonly dispid 103;
    property CUIT: WideString dispid 206;
    property XMLRequest: WideString readonly dispid 201;
    property XMLResponse: WideString readonly dispid 202;
    property Proxy: WideString dispid 203;
    property ProxyUserName: WideString dispid 204;
    property ProxyPassword: WideString dispid 205;
    property ProxyEnabled: OLE_CANCELBOOL dispid 207;
    procedure CargarLicencia(const Licencia: WideString); dispid 208;
  end;

// *********************************************************************//
// Interface: Iwsfexv1
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {10891378-BAE5-4F40-AF39-70C54F4E8175}
// *********************************************************************//
  Iwsfexv1 = interface(IDispatch)
    ['{10891378-BAE5-4F40-AF39-70C54F4E8175}']
    procedure AgregaFactura(Id: Double; const Fecha_cbte: WideString; Tipo_cbte: Integer; 
                            Punto_vta: Integer; Cbte_nro: Double; Tipo_expo: Integer; 
                            const Permiso_existente: WideString; Dst_cmp: Integer; 
                            const Cliente: WideString; Cuit_pais_cliente: Double; 
                            const Domicilio_cliente: WideString; const Id_impositivo: WideString; 
                            const Moneda_Id: WideString; Moneda_ctz: Double; 
                            const Obs_comerciales: WideString; Imp_total: Double; 
                            const Obs: WideString; const Forma_pago: WideString; 
                            const Incoterms: WideString; const Incoterms_ds: WideString; 
                            Idioma_cbte: Integer; const Fecha_pago: WideString); safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    function Autorizar: OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function AutorizarRespuesta(out Cae: WideString; out Fch_venc_Cae: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; safecall;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out Cbte_nro: Double; 
                             out Cbte_fecha: WideString): OLE_CANCELBOOL; safecall;
    procedure AgregaPermiso(const Id_permiso: WideString; Dst_merc: Integer); safecall;
    procedure AgregaCompAsoc(Cbte_tipo: Integer; Cbte_punto_vta: Integer; Cbte_nro: Double; 
                             Cbte_cuit: Double); safecall;
    procedure AgregaItem(const Pro_codigo: WideString; const Pro_ds: WideString; Pro_qty: Double; 
                         Pro_umed: Integer; Pro_precio_uni: Double; Pro_total_item: Double; 
                         Pro_bonificacion: Double); safecall;
    function Get_Token: WideString; safecall;
    procedure Set_Token(const Value: WideString); safecall;
    function Get_Sign: WideString; safecall;
    procedure Set_Sign(const Value: WideString); safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; safecall;
    function Get_SFLastCMP: Double; safecall;
    function Get_SFCAE: WideString; safecall;
    function Get_SFVencimiento: WideString; safecall;
    function Get_SFResultado: WideString; safecall;
    function Get_SFReproceso: WideString; safecall;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; safecall;
    function Get_SFCmpConsultarCAE: WideString; safecall;
    function Get_SFCmpConsultarVencimiento: WideString; safecall;
    function UltimoIdTrans(out Resultado: Double): OLE_CANCELBOOL; safecall;
    function AutorizarRespuestaObs: WideString; safecall;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; safecall;
    function SFUltimoIdTrans: OLE_CANCELBOOL; safecall;
    function Get_SFLastId: Double; safecall;
    function ParamGetCotizacion(const MonId: WideString; var MonCtz: Double; 
                                var MonFecha: WideString): OLE_CANCELBOOL; safecall;
    function Get_Proxy: WideString; safecall;
    procedure Set_Proxy(const Value: WideString); safecall;
    function Get_ProxyUserName: WideString; safecall;
    procedure Set_ProxyUserName(const Value: WideString); safecall;
    function Get_ProxyPassword: WideString; safecall;
    procedure Set_ProxyPassword(const Value: WideString); safecall;
    function Get_ProxyEnabled: OLE_CANCELBOOL; safecall;
    procedure Set_ProxyEnabled(Value: OLE_CANCELBOOL); safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    function ParamGetPtosVenta(var Resultado: WideString): OLE_CANCELBOOL; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property URL: WideString read Get_URL write Set_URL;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property Token: WideString read Get_Token write Set_Token;
    property Sign: WideString read Get_Sign write Set_Sign;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property SFLastCMP: Double read Get_SFLastCMP;
    property SFCAE: WideString read Get_SFCAE;
    property SFVencimiento: WideString read Get_SFVencimiento;
    property SFResultado: WideString read Get_SFResultado;
    property SFReproceso: WideString read Get_SFReproceso;
    property SFCmpConsultarCAE: WideString read Get_SFCmpConsultarCAE;
    property SFCmpConsultarVencimiento: WideString read Get_SFCmpConsultarVencimiento;
    property SFLastId: Double read Get_SFLastId;
    property Proxy: WideString read Get_Proxy write Set_Proxy;
    property ProxyUserName: WideString read Get_ProxyUserName write Set_ProxyUserName;
    property ProxyPassword: WideString read Get_ProxyPassword write Set_ProxyPassword;
    property ProxyEnabled: OLE_CANCELBOOL read Get_ProxyEnabled write Set_ProxyEnabled;
  end;

// *********************************************************************//
// DispIntf:  Iwsfexv1Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {10891378-BAE5-4F40-AF39-70C54F4E8175}
// *********************************************************************//
  Iwsfexv1Disp = dispinterface
    ['{10891378-BAE5-4F40-AF39-70C54F4E8175}']
    procedure AgregaFactura(Id: Double; const Fecha_cbte: WideString; Tipo_cbte: Integer; 
                            Punto_vta: Integer; Cbte_nro: Double; Tipo_expo: Integer; 
                            const Permiso_existente: WideString; Dst_cmp: Integer; 
                            const Cliente: WideString; Cuit_pais_cliente: Double; 
                            const Domicilio_cliente: WideString; const Id_impositivo: WideString; 
                            const Moneda_Id: WideString; Moneda_ctz: Double; 
                            const Obs_comerciales: WideString; Imp_total: Double; 
                            const Obs: WideString; const Forma_pago: WideString; 
                            const Incoterms: WideString; const Incoterms_ds: WideString; 
                            Idioma_cbte: Integer; const Fecha_pago: WideString); dispid 101;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 102;
    function Autorizar: OLE_CANCELBOOL; dispid 103;
    property ErrorCode: Integer readonly dispid 104;
    property ErrorDesc: WideString readonly dispid 105;
    property URL: WideString dispid 107;
    property CUIT: Double dispid 108;
    function AutorizarRespuesta(out Cae: WideString; out Fch_venc_Cae: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; dispid 110;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out Cbte_nro: Double; 
                             out Cbte_fecha: WideString): OLE_CANCELBOOL; dispid 111;
    procedure AgregaPermiso(const Id_permiso: WideString; Dst_merc: Integer); dispid 112;
    procedure AgregaCompAsoc(Cbte_tipo: Integer; Cbte_punto_vta: Integer; Cbte_nro: Double; 
                             Cbte_cuit: Double); dispid 113;
    procedure AgregaItem(const Pro_codigo: WideString; const Pro_ds: WideString; Pro_qty: Double; 
                         Pro_umed: Integer; Pro_precio_uni: Double; Pro_total_item: Double; 
                         Pro_bonificacion: Double); dispid 114;
    property Token: WideString dispid 201;
    property Sign: WideString dispid 202;
    property XMLRequest: WideString readonly dispid 203;
    property XMLResponse: WideString readonly dispid 204;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; dispid 219;
    property SFLastCMP: Double readonly dispid 220;
    property SFCAE: WideString readonly dispid 216;
    property SFVencimiento: WideString readonly dispid 217;
    property SFResultado: WideString readonly dispid 218;
    property SFReproceso: WideString readonly dispid 224;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; dispid 221;
    property SFCmpConsultarCAE: WideString readonly dispid 222;
    property SFCmpConsultarVencimiento: WideString readonly dispid 225;
    function UltimoIdTrans(out Resultado: Double): OLE_CANCELBOOL; dispid 115;
    function AutorizarRespuestaObs: WideString; dispid 205;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; dispid 228;
    function SFUltimoIdTrans: OLE_CANCELBOOL; dispid 230;
    property SFLastId: Double readonly dispid 229;
    function ParamGetCotizacion(const MonId: WideString; var MonCtz: Double; 
                                var MonFecha: WideString): OLE_CANCELBOOL; dispid 206;
    property Proxy: WideString dispid 207;
    property ProxyUserName: WideString dispid 208;
    property ProxyPassword: WideString dispid 209;
    property ProxyEnabled: OLE_CANCELBOOL dispid 210;
    procedure CargarLicencia(const Licencia: WideString); dispid 211;
    function ParamGetPtosVenta(var Resultado: WideString): OLE_CANCELBOOL; dispid 212;
  end;

// *********************************************************************//
// Interface: Iwsfev1
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E0A95BBC-E328-4AA6-84E2-405C10AD41A2}
// *********************************************************************//
  Iwsfev1 = interface(IDispatch)
    ['{E0A95BBC-E328-4AA6-84E2-405C10AD41A2}']
    procedure AgregaFactura(Concepto: Integer; DocTipo: Integer; DocNro: Double; Cbtedesde: Double; 
                            Cbtehasta: Double; const CbteFch: WideString; Imptotal: Double; 
                            ImpTotalConc: Double; ImpNeto: Double; ImpOpEx: Double; 
                            const FechaServDesde: WideString; const FechaServHasta: WideString; 
                            const FechaVencPago: WideString; const MonId: WideString; 
                            MonCotiz: Double); safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    function Autorizar(ptoVenta: Integer; CbteTipo: Integer): OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    procedure Reset; safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_AutorizarRespCount: Integer; safecall;
    function AutorizarRespuesta(Indice: Integer; out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; safecall;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out cmp: Double): OLE_CANCELBOOL; safecall;
    function RecuperaQTYRequest(qty: Integer): OLE_CANCELBOOL; safecall;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; safecall;
    function Dummy(out appserver: WideString; out authserver: WideString; out dbserver: WideString): OLE_CANCELBOOL; safecall;
    function CAEASolicitar(Periodo: Integer; Orden: Integer; out Cae: WideString; 
                           out FchVigDesde: WideString; out FchVigHasta: WideString; 
                           out FchTopeInf: WideString; out FchProceso: WideString): OLE_CANCELBOOL; safecall;
    function AutorizarRespuestaObs(Indice: Integer): WideString; safecall;
    function CAEAConsultar(Periodo: Integer; Orden: Integer; out Cae: WideString; 
                           out FchVigDesde: WideString; out FchVigHasta: WideString; 
                           out FchTopeInf: WideString; out FchProceso: WideString): OLE_CANCELBOOL; safecall;
    function CAEASinMovimientoInformar(PtoVta: Integer; const CAEA: WideString; 
                                       out Resultado: WideString): OLE_CANCELBOOL; safecall;
    function CAEASinMovimientoConsultar(PtoVta: Integer; const CAEA: WideString; 
                                        out Resultado: WideString): OLE_CANCELBOOL; safecall;
    function ParamGetCotizacion(const MonId: WideString; out MonCotiz: Double; 
                                out FchCotiz: WideString): OLE_CANCELBOOL; safecall;
    function ParamGetTiposConcepto(out Resultado: WideString): OLE_CANCELBOOL; safecall;
    procedure AgregaTributo(Id: Integer; const Desc: WideString; BaseImp: Double; Alic: Double; 
                            Importe: Double); safecall;
    procedure AgregaIVA(Id: Integer; BaseImp: Double; Importe: Double); safecall;
    procedure AgregaCompAsoc(Tipo: Integer; PtoVta: Integer; Nro: Double; CUIT: Double; 
                             const CbteFch: WideString); safecall;
    procedure AgregaOpcional(const Id: WideString; const Valor: WideString); safecall;
    function ParamGetTiposMonedas(out Resultado: WideString): OLE_CANCELBOOL; safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function Get_Token: WideString; safecall;
    procedure Set_Token(const Value: WideString); safecall;
    function Get_Sign: WideString; safecall;
    procedure Set_Sign(const Value: WideString); safecall;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; safecall;
    function Get_SFLastCMP: Double; safecall;
    function Get_SFCAE(Indice: Integer): WideString; safecall;
    function Get_SFVencimiento(Indice: Integer): WideString; safecall;
    function Get_SFResultado(Indice: Integer): WideString; safecall;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; safecall;
    function Get_SFCmpConsultarCAE: WideString; safecall;
    function Get_SFCmpConsultarVencimiento: WideString; safecall;
    function CAEAInformar(ptoVenta: Integer; CbteTipo: Integer; const Cae: WideString): OLE_CANCELBOOL; safecall;
    function AutorizarRespuestaObsCode(Indice: Integer): WideString; safecall;
    function Get_Proxy: WideString; safecall;
    procedure Set_Proxy(const Value: WideString); safecall;
    function Get_ProxyUserName: WideString; safecall;
    procedure Set_ProxyUserName(const Value: WideString); safecall;
    function Get_ProxyPassword: WideString; safecall;
    procedure Set_ProxyPassword(const Value: WideString); safecall;
    function Get_ProxyEnabled: OLE_CANCELBOOL; safecall;
    procedure Set_ProxyEnabled(Value: OLE_CANCELBOOL); safecall;
    function ParamGetTiposDoc(out Resultado: WideString): OLE_CANCELBOOL; safecall;
    function ParamGetTiposCbte(out Resultado: WideString): OLE_CANCELBOOL; safecall;
    procedure LogTransaction(const RequestFilename: WideString; const ResponseFilename: WideString); safecall;
    function CmpConsultarEx(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                            const cbte_info_result: IComprobante): OLE_CANCELBOOL; safecall;
    function Get_CmpConsultarCbte: IComprobante; safecall;
    procedure AgregaComprador(DocTipo: Integer; DocNro: Double; Porcentaje: Double); safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    function Get_Path: WideString; safecall;
    function ParamGetPtosVenta(out Resultado: WideString): OLE_CANCELBOOL; safecall;
    procedure PeriodoAsoc(const FchDesde: WideString; const FchHasta: WideString); safecall;
    procedure CAEACbteFchHsGen(const CbteFchHsGen: WideString); safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property URL: WideString read Get_URL write Set_URL;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property AutorizarRespCount: Integer read Get_AutorizarRespCount;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property Token: WideString read Get_Token write Set_Token;
    property Sign: WideString read Get_Sign write Set_Sign;
    property SFLastCMP: Double read Get_SFLastCMP;
    property SFCAE[Indice: Integer]: WideString read Get_SFCAE;
    property SFVencimiento[Indice: Integer]: WideString read Get_SFVencimiento;
    property SFResultado[Indice: Integer]: WideString read Get_SFResultado;
    property SFCmpConsultarCAE: WideString read Get_SFCmpConsultarCAE;
    property SFCmpConsultarVencimiento: WideString read Get_SFCmpConsultarVencimiento;
    property Proxy: WideString read Get_Proxy write Set_Proxy;
    property ProxyUserName: WideString read Get_ProxyUserName write Set_ProxyUserName;
    property ProxyPassword: WideString read Get_ProxyPassword write Set_ProxyPassword;
    property ProxyEnabled: OLE_CANCELBOOL read Get_ProxyEnabled write Set_ProxyEnabled;
    property CmpConsultarCbte: IComprobante read Get_CmpConsultarCbte;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property Path: WideString read Get_Path;
  end;

// *********************************************************************//
// DispIntf:  Iwsfev1Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E0A95BBC-E328-4AA6-84E2-405C10AD41A2}
// *********************************************************************//
  Iwsfev1Disp = dispinterface
    ['{E0A95BBC-E328-4AA6-84E2-405C10AD41A2}']
    procedure AgregaFactura(Concepto: Integer; DocTipo: Integer; DocNro: Double; Cbtedesde: Double; 
                            Cbtehasta: Double; const CbteFch: WideString; Imptotal: Double; 
                            ImpTotalConc: Double; ImpNeto: Double; ImpOpEx: Double; 
                            const FechaServDesde: WideString; const FechaServHasta: WideString; 
                            const FechaVencPago: WideString; const MonId: WideString; 
                            MonCotiz: Double); dispid 201;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 202;
    function Autorizar(ptoVenta: Integer; CbteTipo: Integer): OLE_CANCELBOOL; dispid 203;
    property ErrorCode: Integer readonly dispid 204;
    property ErrorDesc: WideString readonly dispid 205;
    procedure Reset; dispid 206;
    property URL: WideString dispid 207;
    property CUIT: Double dispid 208;
    property AutorizarRespCount: Integer readonly dispid 209;
    function AutorizarRespuesta(Indice: Integer; out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; dispid 210;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out cmp: Double): OLE_CANCELBOOL; dispid 211;
    function RecuperaQTYRequest(qty: Integer): OLE_CANCELBOOL; dispid 212;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; dispid 213;
    function Dummy(out appserver: WideString; out authserver: WideString; out dbserver: WideString): OLE_CANCELBOOL; dispid 214;
    function CAEASolicitar(Periodo: Integer; Orden: Integer; out Cae: WideString; 
                           out FchVigDesde: WideString; out FchVigHasta: WideString; 
                           out FchTopeInf: WideString; out FchProceso: WideString): OLE_CANCELBOOL; dispid 215;
    function AutorizarRespuestaObs(Indice: Integer): WideString; dispid 216;
    function CAEAConsultar(Periodo: Integer; Orden: Integer; out Cae: WideString; 
                           out FchVigDesde: WideString; out FchVigHasta: WideString; 
                           out FchTopeInf: WideString; out FchProceso: WideString): OLE_CANCELBOOL; dispid 217;
    function CAEASinMovimientoInformar(PtoVta: Integer; const CAEA: WideString; 
                                       out Resultado: WideString): OLE_CANCELBOOL; dispid 218;
    function CAEASinMovimientoConsultar(PtoVta: Integer; const CAEA: WideString; 
                                        out Resultado: WideString): OLE_CANCELBOOL; dispid 219;
    function ParamGetCotizacion(const MonId: WideString; out MonCotiz: Double; 
                                out FchCotiz: WideString): OLE_CANCELBOOL; dispid 220;
    function ParamGetTiposConcepto(out Resultado: WideString): OLE_CANCELBOOL; dispid 221;
    procedure AgregaTributo(Id: Integer; const Desc: WideString; BaseImp: Double; Alic: Double; 
                            Importe: Double); dispid 222;
    procedure AgregaIVA(Id: Integer; BaseImp: Double; Importe: Double); dispid 223;
    procedure AgregaCompAsoc(Tipo: Integer; PtoVta: Integer; Nro: Double; CUIT: Double; 
                             const CbteFch: WideString); dispid 224;
    procedure AgregaOpcional(const Id: WideString; const Valor: WideString); dispid 225;
    function ParamGetTiposMonedas(out Resultado: WideString): OLE_CANCELBOOL; dispid 226;
    property XMLRequest: WideString readonly dispid 227;
    property XMLResponse: WideString readonly dispid 228;
    property Token: WideString dispid 229;
    property Sign: WideString dispid 230;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; dispid 231;
    property SFLastCMP: Double readonly dispid 232;
    property SFCAE[Indice: Integer]: WideString readonly dispid 233;
    property SFVencimiento[Indice: Integer]: WideString readonly dispid 234;
    property SFResultado[Indice: Integer]: WideString readonly dispid 235;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; dispid 236;
    property SFCmpConsultarCAE: WideString readonly dispid 237;
    property SFCmpConsultarVencimiento: WideString readonly dispid 238;
    function CAEAInformar(ptoVenta: Integer; CbteTipo: Integer; const Cae: WideString): OLE_CANCELBOOL; dispid 239;
    function AutorizarRespuestaObsCode(Indice: Integer): WideString; dispid 240;
    property Proxy: WideString dispid 241;
    property ProxyUserName: WideString dispid 242;
    property ProxyPassword: WideString dispid 243;
    property ProxyEnabled: OLE_CANCELBOOL dispid 244;
    function ParamGetTiposDoc(out Resultado: WideString): OLE_CANCELBOOL; dispid 245;
    function ParamGetTiposCbte(out Resultado: WideString): OLE_CANCELBOOL; dispid 246;
    procedure LogTransaction(const RequestFilename: WideString; const ResponseFilename: WideString); dispid 247;
    function CmpConsultarEx(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                            const cbte_info_result: IComprobante): OLE_CANCELBOOL; dispid 249;
    property CmpConsultarCbte: IComprobante readonly dispid 248;
    procedure AgregaComprador(DocTipo: Integer; DocNro: Double; Porcentaje: Double); dispid 250;
    property Depurar: OLE_CANCELBOOL dispid 251;
    procedure CargarLicencia(const Licencia: WideString); dispid 252;
    property Path: WideString readonly dispid 253;
    function ParamGetPtosVenta(out Resultado: WideString): OLE_CANCELBOOL; dispid 254;
    procedure PeriodoAsoc(const FchDesde: WideString; const FchHasta: WideString); dispid 255;
    procedure CAEACbteFchHsGen(const CbteFchHsGen: WideString); dispid 256;
  end;

// *********************************************************************//
// Interface: Iwsbfev1
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A5C9683D-3D72-4392-AD49-A4DFB83D8C63}
// *********************************************************************//
  Iwsbfev1 = interface(IDispatch)
    ['{A5C9683D-3D72-4392-AD49-A4DFB83D8C63}']
    procedure AgregaFactura(Id: Double; tipo_doc: Integer; nro_doc: Double; Zona: Integer; 
                            Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                            Imp_total: Double; imp_tot_conc: Double; imp_neto: Double; 
                            impto_liq: Double; impto_liq_rni: Double; imp_op_ex: Double; 
                            Imp_perc: Double; Imp_iibb: Double; Imp_perc_mun: Double; 
                            Imp_internos: Double; const Imp_moneda_Id: WideString; 
                            Imp_moneda_ctz: Double; const Fecha_cbte: WideString; 
                            const Fecha_vto_pago: WideString); safecall;
    procedure AgregaOpcional(const Id: WideString; const Valor: WideString); safecall;
    procedure AgregaItem(const Pro_codigo_ncm: WideString; const Pro_codigo_sec: WideString; 
                         const Pro_ds: WideString; Pro_qty: Double; Pro_umed: Integer; 
                         Pro_precio_uni: Double; Imp_bonif: Double; Imp_total: Double; 
                         Iva_id: Integer); safecall;
    function Autorizar: OLE_CANCELBOOL; safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_Token: WideString; safecall;
    procedure Set_Token(const Value: WideString); safecall;
    function Get_Sign: WideString; safecall;
    procedure Set_Sign(const Value: WideString); safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    procedure Reset; safecall;
    function RecuperaLastCMP(Pto_venta: Integer; Tipo_cbte: Integer; out Cbte_nro: Double; 
                             out Cbte_fecha: WideString): OLE_CANCELBOOL; safecall;
    function SFRecuperaLastCMP(Pto_venta: Integer; Tipo_cbte: Integer): OLE_CANCELBOOL; safecall;
    function Get_SFLastCMP: Double; safecall;
    function Get_SFLastFecha: WideString; safecall;
    function RecuperaLastID(out Id: Double): OLE_CANCELBOOL; safecall;
    function SFRecuperaLastID: OLE_CANCELBOOL; safecall;
    function Get_SFLastId: Double; safecall;
    function AutorizarRespuesta(out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; safecall;
    function Get_SFCAE: WideString; safecall;
    function Get_SFVencimiento: WideString; safecall;
    function Get_SFResultado: WideString; safecall;
    function Get_SFReproceso: WideString; safecall;
    function AutorizarRespuestaObs: WideString; safecall;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; safecall;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Double): OLE_CANCELBOOL; safecall;
    function Get_SFCmpConsultarCAE: WideString; safecall;
    function Get_SFCmpConsultarVencimiento: WideString; safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function Get_Proxy: WideString; safecall;
    procedure Set_Proxy(const Value: WideString); safecall;
    function Get_ProxyUserName: WideString; safecall;
    procedure Set_ProxyUserName(const Value: WideString); safecall;
    function Get_ProxyPassword: WideString; safecall;
    procedure Set_ProxyPassword(const Value: WideString); safecall;
    function Get_ProxyEnabled: OLE_CANCELBOOL; safecall;
    procedure Set_ProxyEnabled(Value: OLE_CANCELBOOL); safecall;
    function ParamGetZonas(out Zonas: WideString): OLE_CANCELBOOL; safecall;
    procedure AgregaCompAsoc(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Double; 
                             CUIT: Double; const Fecha_cbte: WideString); safecall;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property URL: WideString read Get_URL write Set_URL;
    property Token: WideString read Get_Token write Set_Token;
    property Sign: WideString read Get_Sign write Set_Sign;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property SFLastCMP: Double read Get_SFLastCMP;
    property SFLastFecha: WideString read Get_SFLastFecha;
    property SFLastId: Double read Get_SFLastId;
    property SFCAE: WideString read Get_SFCAE;
    property SFVencimiento: WideString read Get_SFVencimiento;
    property SFResultado: WideString read Get_SFResultado;
    property SFReproceso: WideString read Get_SFReproceso;
    property SFCmpConsultarCAE: WideString read Get_SFCmpConsultarCAE;
    property SFCmpConsultarVencimiento: WideString read Get_SFCmpConsultarVencimiento;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property Proxy: WideString read Get_Proxy write Set_Proxy;
    property ProxyUserName: WideString read Get_ProxyUserName write Set_ProxyUserName;
    property ProxyPassword: WideString read Get_ProxyPassword write Set_ProxyPassword;
    property ProxyEnabled: OLE_CANCELBOOL read Get_ProxyEnabled write Set_ProxyEnabled;
  end;

// *********************************************************************//
// DispIntf:  Iwsbfev1Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A5C9683D-3D72-4392-AD49-A4DFB83D8C63}
// *********************************************************************//
  Iwsbfev1Disp = dispinterface
    ['{A5C9683D-3D72-4392-AD49-A4DFB83D8C63}']
    procedure AgregaFactura(Id: Double; tipo_doc: Integer; nro_doc: Double; Zona: Integer; 
                            Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                            Imp_total: Double; imp_tot_conc: Double; imp_neto: Double; 
                            impto_liq: Double; impto_liq_rni: Double; imp_op_ex: Double; 
                            Imp_perc: Double; Imp_iibb: Double; Imp_perc_mun: Double; 
                            Imp_internos: Double; const Imp_moneda_Id: WideString; 
                            Imp_moneda_ctz: Double; const Fecha_cbte: WideString; 
                            const Fecha_vto_pago: WideString); dispid 201;
    procedure AgregaOpcional(const Id: WideString; const Valor: WideString); dispid 202;
    procedure AgregaItem(const Pro_codigo_ncm: WideString; const Pro_codigo_sec: WideString; 
                         const Pro_ds: WideString; Pro_qty: Double; Pro_umed: Integer; 
                         Pro_precio_uni: Double; Imp_bonif: Double; Imp_total: Double; 
                         Iva_id: Integer); dispid 203;
    function Autorizar: OLE_CANCELBOOL; dispid 204;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 205;
    property CUIT: Double dispid 206;
    property URL: WideString dispid 207;
    property Token: WideString dispid 208;
    property Sign: WideString dispid 209;
    property ErrorCode: Integer readonly dispid 210;
    property ErrorDesc: WideString readonly dispid 211;
    procedure Reset; dispid 212;
    function RecuperaLastCMP(Pto_venta: Integer; Tipo_cbte: Integer; out Cbte_nro: Double; 
                             out Cbte_fecha: WideString): OLE_CANCELBOOL; dispid 213;
    function SFRecuperaLastCMP(Pto_venta: Integer; Tipo_cbte: Integer): OLE_CANCELBOOL; dispid 214;
    property SFLastCMP: Double readonly dispid 215;
    property SFLastFecha: WideString readonly dispid 216;
    function RecuperaLastID(out Id: Double): OLE_CANCELBOOL; dispid 217;
    function SFRecuperaLastID: OLE_CANCELBOOL; dispid 219;
    property SFLastId: Double readonly dispid 218;
    function AutorizarRespuesta(out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString; out Reproceso: WideString): OLE_CANCELBOOL; dispid 220;
    property SFCAE: WideString readonly dispid 222;
    property SFVencimiento: WideString readonly dispid 223;
    property SFResultado: WideString readonly dispid 224;
    property SFReproceso: WideString readonly dispid 225;
    function AutorizarRespuestaObs: WideString; dispid 226;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; dispid 227;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Double): OLE_CANCELBOOL; dispid 228;
    property SFCmpConsultarCAE: WideString readonly dispid 229;
    property SFCmpConsultarVencimiento: WideString readonly dispid 230;
    property XMLRequest: WideString readonly dispid 231;
    property XMLResponse: WideString readonly dispid 232;
    property Proxy: WideString dispid 221;
    property ProxyUserName: WideString dispid 233;
    property ProxyPassword: WideString dispid 234;
    property ProxyEnabled: OLE_CANCELBOOL dispid 235;
    function ParamGetZonas(out Zonas: WideString): OLE_CANCELBOOL; dispid 236;
    procedure AgregaCompAsoc(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Double; 
                             CUIT: Double; const Fecha_cbte: WideString); dispid 237;
  end;

// *********************************************************************//
// Interface: Iwsmtxca
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C297BD2B-A528-446B-BF55-FAF195383E0E}
// *********************************************************************//
  Iwsmtxca = interface(IDispatch)
    ['{C297BD2B-A528-446B-BF55-FAF195383E0E}']
    procedure AgregaFactura(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                            numeroComprobante: Double; const fechaEmision: WideString; 
                            codigoTipoDocumento: Integer; numeroDocumento: Double; 
                            importeGravado: Double; importeNoGravado: Double; 
                            importeExento: Double; importeSubtotal: Double; 
                            importeOtrosTributos: Double; importeTotal: Double; 
                            const codigoMoneda: WideString; cotizacionMoneda: Double; 
                            const observaciones: WideString; codigoConcepto: Integer; 
                            const fechaServicioDesde: WideString; 
                            const fechaServicioHasta: WideString; 
                            const fechaVencimientoPago: WideString); safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    procedure AgregaTributo(Id: Integer; const Desc: WideString; BaseImp: Double; Importe: Double); safecall;
    procedure AgregaIVA(codigo: Integer; Importe: Double); safecall;
    procedure AgregaCompAsoc(Tipo: Integer; PtoVta: Integer; Nro: Double; CUIT: Double; 
                             const fechaEmision: WideString); safecall;
    procedure AgregaItem(unidadesMtx: Integer; const codigoMtx: WideString; 
                         const codigo: WideString; const descripcion: WideString; cantidad: Double; 
                         codigoUnidadMedida: Integer; precioUnitario: Double; 
                         importeBonificacion: Double; codigoCondicionIVA: Integer; 
                         importeIVA: Double; importeItem: Double); safecall;
    function Autorizar: OLE_CANCELBOOL; safecall;
    function AutorizarRespuesta(out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString): OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out cmp: Double): OLE_CANCELBOOL; safecall;
    function AutorizarRespuestaObs: WideString; safecall;
    function Get_Token: WideString; safecall;
    procedure Set_Token(const Value: WideString); safecall;
    function Get_Sign: WideString; safecall;
    procedure Set_Sign(const Value: WideString); safecall;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; safecall;
    function Get_SFCmpConsultarCAE: WideString; safecall;
    function Get_SFCmpConsultarVencimiento: WideString; safecall;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; safecall;
    function Get_SFLastCMP: Double; safecall;
    function Get_SFCAE: WideString; safecall;
    function Get_SFVencimiento: WideString; safecall;
    function Get_SFResultado: WideString; safecall;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    procedure AgregaDatoAdicional(T: Integer; const C1: WideString; const C2: WideString; 
                                  const C3: WideString; const C4: WideString; const C5: WideString; 
                                  const C6: WideString); safecall;
    procedure PeriodoCompAsoc(const fechaDesde: WideString; const fechaHasta: WideString); safecall;
    procedure AgregaComprador(codigoTipoDocumento: Integer; numeroDocumento: Double; 
                              Porcentaje: Double); safecall;
    function ParamGetPtosVenta(var Resultado: WideString): OLE_CANCELBOOL; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property URL: WideString read Get_URL write Set_URL;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property Token: WideString read Get_Token write Set_Token;
    property Sign: WideString read Get_Sign write Set_Sign;
    property SFCmpConsultarCAE: WideString read Get_SFCmpConsultarCAE;
    property SFCmpConsultarVencimiento: WideString read Get_SFCmpConsultarVencimiento;
    property SFLastCMP: Double read Get_SFLastCMP;
    property SFCAE: WideString read Get_SFCAE;
    property SFVencimiento: WideString read Get_SFVencimiento;
    property SFResultado: WideString read Get_SFResultado;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
  end;

// *********************************************************************//
// DispIntf:  IwsmtxcaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C297BD2B-A528-446B-BF55-FAF195383E0E}
// *********************************************************************//
  IwsmtxcaDisp = dispinterface
    ['{C297BD2B-A528-446B-BF55-FAF195383E0E}']
    procedure AgregaFactura(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                            numeroComprobante: Double; const fechaEmision: WideString; 
                            codigoTipoDocumento: Integer; numeroDocumento: Double; 
                            importeGravado: Double; importeNoGravado: Double; 
                            importeExento: Double; importeSubtotal: Double; 
                            importeOtrosTributos: Double; importeTotal: Double; 
                            const codigoMoneda: WideString; cotizacionMoneda: Double; 
                            const observaciones: WideString; codigoConcepto: Integer; 
                            const fechaServicioDesde: WideString; 
                            const fechaServicioHasta: WideString; 
                            const fechaVencimientoPago: WideString); dispid 101;
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 102;
    procedure AgregaTributo(Id: Integer; const Desc: WideString; BaseImp: Double; Importe: Double); dispid 208;
    procedure AgregaIVA(codigo: Integer; Importe: Double); dispid 209;
    procedure AgregaCompAsoc(Tipo: Integer; PtoVta: Integer; Nro: Double; CUIT: Double; 
                             const fechaEmision: WideString); dispid 210;
    procedure AgregaItem(unidadesMtx: Integer; const codigoMtx: WideString; 
                         const codigo: WideString; const descripcion: WideString; cantidad: Double; 
                         codigoUnidadMedida: Integer; precioUnitario: Double; 
                         importeBonificacion: Double; codigoCondicionIVA: Integer; 
                         importeIVA: Double; importeItem: Double); dispid 201;
    function Autorizar: OLE_CANCELBOOL; dispid 103;
    function AutorizarRespuesta(out Cae: WideString; out Vencimiento: WideString; 
                                out Resultado: WideString): OLE_CANCELBOOL; dispid 110;
    property ErrorCode: Integer readonly dispid 104;
    property ErrorDesc: WideString readonly dispid 105;
    property URL: WideString dispid 107;
    property CUIT: Double dispid 108;
    function RecuperaLastCMP(PtoVta: Integer; TipoComp: Integer; out cmp: Double): OLE_CANCELBOOL; dispid 112;
    function AutorizarRespuestaObs: WideString; dispid 202;
    property Token: WideString dispid 203;
    property Sign: WideString dispid 204;
    function SFCmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double): OLE_CANCELBOOL; dispid 218;
    property SFCmpConsultarCAE: WideString readonly dispid 219;
    property SFCmpConsultarVencimiento: WideString readonly dispid 220;
    function SFRecuperaLastCMP(PtoVta: Integer; TipoComp: Integer): OLE_CANCELBOOL; dispid 221;
    property SFLastCMP: Double readonly dispid 222;
    property SFCAE: WideString readonly dispid 223;
    property SFVencimiento: WideString readonly dispid 224;
    property SFResultado: WideString readonly dispid 225;
    function CmpConsultar(Tipo_cbte: Integer; Punto_vta: Integer; cbt_nro: Double; 
                          out Cae: WideString; out Vencimiento: WideString): OLE_CANCELBOOL; dispid 228;
    property XMLRequest: WideString readonly dispid 229;
    property XMLResponse: WideString readonly dispid 230;
    property Depurar: OLE_CANCELBOOL dispid 205;
    procedure AgregaDatoAdicional(T: Integer; const C1: WideString; const C2: WideString; 
                                  const C3: WideString; const C4: WideString; const C5: WideString; 
                                  const C6: WideString); dispid 206;
    procedure PeriodoCompAsoc(const fechaDesde: WideString; const fechaHasta: WideString); dispid 207;
    procedure AgregaComprador(codigoTipoDocumento: Integer; numeroDocumento: Double; 
                              Porcentaje: Double); dispid 212;
    function ParamGetPtosVenta(var Resultado: WideString): OLE_CANCELBOOL; dispid 211;
  end;

// *********************************************************************//
// Interface: Iwsseg
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B1E85685-67E8-4B99-B8B6-85A6138E4DD0}
// *********************************************************************//
  Iwsseg = interface(IDispatch)
    ['{B1E85685-67E8-4B99-B8B6-85A6138E4DD0}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    procedure AgregaFactura(Id: Integer; tipo_doc: Integer; nro_doc: Double; Tipo_cbte: Integer; 
                            Punto_vta: Integer; Cbte_nro: Integer; Imp_total: Double; 
                            imp_tot_conc: Double; imp_neto: Double; impto_liq: Double; 
                            impto_liq_rni: Double; imp_op_ex: Double; Imp_perc: Double; 
                            Imp_iibb: Double; Imp_perc_mun: Double; Imp_internos: Double; 
                            const Imp_moneda_Id: WideString; Imp_moneda_ctz: Double; 
                            Imp_otrib_prov: Double; const Fecha_cbte: WideString); safecall;
    procedure AgregaItem(const Poliza: WideString; const Endoso: WideString; const Ds: WideString; 
                         qty: Double; Precio_uni: Double; Imp_bonif: Double; Imp_total: Double; 
                         Imp_valor_aseg: Double; const Imp_moneda_vaseg: WideString; Iva_id: Integer); safecall;
    function Autorizar: OLE_CANCELBOOL; safecall;
    function Get_RespuestaAutorizarCAE: WideString; safecall;
    function Get_RespuestaAutorizarVencimiento: WideString; safecall;
    function Get_RespuestaAutorizarResultado: WideString; safecall;
    function Get_RespuestaAutorizarReproceso: WideString; safecall;
    function GetLast_CMP(Pto_venta: Integer; Tipo_cbte: Integer): OLE_CANCELBOOL; safecall;
    function Get_RespuestaGetLast_CMPNro: Integer; safecall;
    function Get_RespuestaGetLast_CMPFecha: WideString; safecall;
    function GetLast_ID: OLE_CANCELBOOL; safecall;
    function Get_RespuestaGetLast_IDId: Integer; safecall;
    function GetCMP(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Integer): OLE_CANCELBOOL; safecall;
    function Get_RespuestaAutorizarObs: WideString; safecall;
    procedure LogTransaction(const RequestFilename: WideString; const ResponseFilename: WideString); safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property URL: WideString read Get_URL write Set_URL;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property RespuestaAutorizarCAE: WideString read Get_RespuestaAutorizarCAE;
    property RespuestaAutorizarVencimiento: WideString read Get_RespuestaAutorizarVencimiento;
    property RespuestaAutorizarResultado: WideString read Get_RespuestaAutorizarResultado;
    property RespuestaAutorizarReproceso: WideString read Get_RespuestaAutorizarReproceso;
    property RespuestaGetLast_CMPNro: Integer read Get_RespuestaGetLast_CMPNro;
    property RespuestaGetLast_CMPFecha: WideString read Get_RespuestaGetLast_CMPFecha;
    property RespuestaGetLast_IDId: Integer read Get_RespuestaGetLast_IDId;
    property RespuestaAutorizarObs: WideString read Get_RespuestaAutorizarObs;
  end;

// *********************************************************************//
// DispIntf:  IwssegDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B1E85685-67E8-4B99-B8B6-85A6138E4DD0}
// *********************************************************************//
  IwssegDisp = dispinterface
    ['{B1E85685-67E8-4B99-B8B6-85A6138E4DD0}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 201;
    property ErrorCode: Integer readonly dispid 202;
    property ErrorDesc: WideString readonly dispid 203;
    property URL: WideString dispid 205;
    property CUIT: Double dispid 204;
    property XMLRequest: WideString readonly dispid 206;
    property XMLResponse: WideString readonly dispid 207;
    procedure AgregaFactura(Id: Integer; tipo_doc: Integer; nro_doc: Double; Tipo_cbte: Integer; 
                            Punto_vta: Integer; Cbte_nro: Integer; Imp_total: Double; 
                            imp_tot_conc: Double; imp_neto: Double; impto_liq: Double; 
                            impto_liq_rni: Double; imp_op_ex: Double; Imp_perc: Double; 
                            Imp_iibb: Double; Imp_perc_mun: Double; Imp_internos: Double; 
                            const Imp_moneda_Id: WideString; Imp_moneda_ctz: Double; 
                            Imp_otrib_prov: Double; const Fecha_cbte: WideString); dispid 208;
    procedure AgregaItem(const Poliza: WideString; const Endoso: WideString; const Ds: WideString; 
                         qty: Double; Precio_uni: Double; Imp_bonif: Double; Imp_total: Double; 
                         Imp_valor_aseg: Double; const Imp_moneda_vaseg: WideString; Iva_id: Integer); dispid 209;
    function Autorizar: OLE_CANCELBOOL; dispid 210;
    property RespuestaAutorizarCAE: WideString readonly dispid 211;
    property RespuestaAutorizarVencimiento: WideString readonly dispid 212;
    property RespuestaAutorizarResultado: WideString readonly dispid 213;
    property RespuestaAutorizarReproceso: WideString readonly dispid 214;
    function GetLast_CMP(Pto_venta: Integer; Tipo_cbte: Integer): OLE_CANCELBOOL; dispid 215;
    property RespuestaGetLast_CMPNro: Integer readonly dispid 216;
    property RespuestaGetLast_CMPFecha: WideString readonly dispid 217;
    function GetLast_ID: OLE_CANCELBOOL; dispid 218;
    property RespuestaGetLast_IDId: Integer readonly dispid 219;
    function GetCMP(Tipo_cbte: Integer; Punto_vta: Integer; Cbte_nro: Integer): OLE_CANCELBOOL; dispid 220;
    property RespuestaAutorizarObs: WideString readonly dispid 221;
    procedure LogTransaction(const RequestFilename: WideString; const ResponseFilename: WideString); dispid 222;
  end;

// *********************************************************************//
// Interface: IwsPadron
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0CEB0878-6393-4701-8C86-2CA793CDCB0D}
// *********************************************************************//
  IwsPadron = interface(IDispatch)
    ['{0CEB0878-6393-4701-8C86-2CA793CDCB0D}']
    function consultar(CUIT: Double; const contribuyenteResult: IContribuyente): OLE_CANCELBOOL; safecall;
    function descargarConstancia(CUIT: Double; const ArchivoDestino: WideString): OLE_CANCELBOOL; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_ModoProduccion: OLE_CANCELBOOL; safecall;
    procedure Set_ModoProduccion(Value: OLE_CANCELBOOL); safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    function sfConsultar(CUIT: Double): OLE_CANCELBOOL; safecall;
    function Get_Contribuyente: IContribuyente; safecall;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property ModoProduccion: OLE_CANCELBOOL read Get_ModoProduccion write Set_ModoProduccion;
    property Contribuyente: IContribuyente read Get_Contribuyente;
  end;

// *********************************************************************//
// DispIntf:  IwsPadronDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0CEB0878-6393-4701-8C86-2CA793CDCB0D}
// *********************************************************************//
  IwsPadronDisp = dispinterface
    ['{0CEB0878-6393-4701-8C86-2CA793CDCB0D}']
    function consultar(CUIT: Double; const contribuyenteResult: IContribuyente): OLE_CANCELBOOL; dispid 201;
    function descargarConstancia(CUIT: Double; const ArchivoDestino: WideString): OLE_CANCELBOOL; dispid 202;
    property ErrorDesc: WideString readonly dispid 206;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; dispid 203;
    property CUIT: Double dispid 204;
    property ModoProduccion: OLE_CANCELBOOL dispid 205;
    procedure CargarLicencia(const Licencia: WideString); dispid 207;
    function sfConsultar(CUIT: Double): OLE_CANCELBOOL; dispid 208;
    property Contribuyente: IContribuyente readonly dispid 209;
  end;

// *********************************************************************//
// Interface: IComprobante
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DC4152DF-68E8-4C5C-804F-22B28CF4C726}
// *********************************************************************//
  IComprobante = interface(IDispatch)
    ['{DC4152DF-68E8-4C5C-804F-22B28CF4C726}']
    function Get_Concepto: Integer; safecall;
    function Get_DocTipo: Integer; safecall;
    function Get_DocNro: Double; safecall;
    function Get_Cbtedesde: Double; safecall;
    function Get_Cbtehasta: Double; safecall;
    function Get_CbteFch: WideString; safecall;
    function Get_Imptotal: Double; safecall;
    function Get_ImpTotConc: Double; safecall;
    function Get_ImpNeto: Double; safecall;
    function Get_ImpOpEx: Double; safecall;
    function Get_ImpTrib: Double; safecall;
    function Get_ImpIVA: Double; safecall;
    function Get_FchServDesde: WideString; safecall;
    function Get_FchServHasta: WideString; safecall;
    function Get_FchVtoPago: WideString; safecall;
    function Get_MonId: WideString; safecall;
    function Get_MonCotiz: Double; safecall;
    function Get_CbtesAsocCount: Integer; safecall;
    function CbtesAsoc(Indice: Integer): ICbteAsoc; safecall;
    function Get_TributosCount: Integer; safecall;
    function Tributos(Indice: Integer): ITributo; safecall;
    function Get_IvaCount: Integer; safecall;
    function Iva(Indice: Integer): IAlicIva; safecall;
    function Get_OpcionalesCount: Integer; safecall;
    function Opcionales(Indice: Integer): IOpcional; safecall;
    function Get_Resultado: WideString; safecall;
    function Get_CodAutorizacion: WideString; safecall;
    function Get_EmisionTipo: WideString; safecall;
    function Get_FchVto: WideString; safecall;
    function Get_FchProceso: WideString; safecall;
    function Get_ObservacionesCount: Integer; safecall;
    function observaciones(Indice: Integer): IObs; safecall;
    function Get_PtoVta: Integer; safecall;
    function Get_CbteTipo: Integer; safecall;
    property Concepto: Integer read Get_Concepto;
    property DocTipo: Integer read Get_DocTipo;
    property DocNro: Double read Get_DocNro;
    property Cbtedesde: Double read Get_Cbtedesde;
    property Cbtehasta: Double read Get_Cbtehasta;
    property CbteFch: WideString read Get_CbteFch;
    property Imptotal: Double read Get_Imptotal;
    property ImpTotConc: Double read Get_ImpTotConc;
    property ImpNeto: Double read Get_ImpNeto;
    property ImpOpEx: Double read Get_ImpOpEx;
    property ImpTrib: Double read Get_ImpTrib;
    property ImpIVA: Double read Get_ImpIVA;
    property FchServDesde: WideString read Get_FchServDesde;
    property FchServHasta: WideString read Get_FchServHasta;
    property FchVtoPago: WideString read Get_FchVtoPago;
    property MonId: WideString read Get_MonId;
    property MonCotiz: Double read Get_MonCotiz;
    property CbtesAsocCount: Integer read Get_CbtesAsocCount;
    property TributosCount: Integer read Get_TributosCount;
    property IvaCount: Integer read Get_IvaCount;
    property OpcionalesCount: Integer read Get_OpcionalesCount;
    property Resultado: WideString read Get_Resultado;
    property CodAutorizacion: WideString read Get_CodAutorizacion;
    property EmisionTipo: WideString read Get_EmisionTipo;
    property FchVto: WideString read Get_FchVto;
    property FchProceso: WideString read Get_FchProceso;
    property ObservacionesCount: Integer read Get_ObservacionesCount;
    property PtoVta: Integer read Get_PtoVta;
    property CbteTipo: Integer read Get_CbteTipo;
  end;

// *********************************************************************//
// DispIntf:  IComprobanteDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DC4152DF-68E8-4C5C-804F-22B28CF4C726}
// *********************************************************************//
  IComprobanteDisp = dispinterface
    ['{DC4152DF-68E8-4C5C-804F-22B28CF4C726}']
    property Concepto: Integer readonly dispid 201;
    property DocTipo: Integer readonly dispid 202;
    property DocNro: Double readonly dispid 203;
    property Cbtedesde: Double readonly dispid 204;
    property Cbtehasta: Double readonly dispid 205;
    property CbteFch: WideString readonly dispid 206;
    property Imptotal: Double readonly dispid 207;
    property ImpTotConc: Double readonly dispid 208;
    property ImpNeto: Double readonly dispid 209;
    property ImpOpEx: Double readonly dispid 210;
    property ImpTrib: Double readonly dispid 211;
    property ImpIVA: Double readonly dispid 212;
    property FchServDesde: WideString readonly dispid 213;
    property FchServHasta: WideString readonly dispid 214;
    property FchVtoPago: WideString readonly dispid 215;
    property MonId: WideString readonly dispid 216;
    property MonCotiz: Double readonly dispid 217;
    property CbtesAsocCount: Integer readonly dispid 218;
    function CbtesAsoc(Indice: Integer): ICbteAsoc; dispid 219;
    property TributosCount: Integer readonly dispid 220;
    function Tributos(Indice: Integer): ITributo; dispid 221;
    property IvaCount: Integer readonly dispid 222;
    function Iva(Indice: Integer): IAlicIva; dispid 223;
    property OpcionalesCount: Integer readonly dispid 224;
    function Opcionales(Indice: Integer): IOpcional; dispid 225;
    property Resultado: WideString readonly dispid 226;
    property CodAutorizacion: WideString readonly dispid 227;
    property EmisionTipo: WideString readonly dispid 228;
    property FchVto: WideString readonly dispid 229;
    property FchProceso: WideString readonly dispid 230;
    property ObservacionesCount: Integer readonly dispid 231;
    function observaciones(Indice: Integer): IObs; dispid 232;
    property PtoVta: Integer readonly dispid 233;
    property CbteTipo: Integer readonly dispid 234;
  end;

// *********************************************************************//
// Interface: ICbteAsoc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {43E44C59-376E-4A27-93D2-ADC712D2BA2E}
// *********************************************************************//
  ICbteAsoc = interface(IDispatch)
    ['{43E44C59-376E-4A27-93D2-ADC712D2BA2E}']
    function Get_Tipo: Integer; safecall;
    function Get_PtoVta: Integer; safecall;
    function Get_Nro: Double; safecall;
    property Tipo: Integer read Get_Tipo;
    property PtoVta: Integer read Get_PtoVta;
    property Nro: Double read Get_Nro;
  end;

// *********************************************************************//
// DispIntf:  ICbteAsocDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {43E44C59-376E-4A27-93D2-ADC712D2BA2E}
// *********************************************************************//
  ICbteAsocDisp = dispinterface
    ['{43E44C59-376E-4A27-93D2-ADC712D2BA2E}']
    property Tipo: Integer readonly dispid 201;
    property PtoVta: Integer readonly dispid 202;
    property Nro: Double readonly dispid 203;
  end;

// *********************************************************************//
// Interface: ITributo
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8C1BE2D0-B8B0-442E-A8D6-8BBBE941DB0C}
// *********************************************************************//
  ITributo = interface(IDispatch)
    ['{8C1BE2D0-B8B0-442E-A8D6-8BBBE941DB0C}']
    function Get_Id: Integer; safecall;
    function Get_Desc: WideString; safecall;
    function Get_BaseImp: Double; safecall;
    function Get_Alic: Double; safecall;
    function Get_Importe: Double; safecall;
    property Id: Integer read Get_Id;
    property Desc: WideString read Get_Desc;
    property BaseImp: Double read Get_BaseImp;
    property Alic: Double read Get_Alic;
    property Importe: Double read Get_Importe;
  end;

// *********************************************************************//
// DispIntf:  ITributoDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8C1BE2D0-B8B0-442E-A8D6-8BBBE941DB0C}
// *********************************************************************//
  ITributoDisp = dispinterface
    ['{8C1BE2D0-B8B0-442E-A8D6-8BBBE941DB0C}']
    property Id: Integer readonly dispid 201;
    property Desc: WideString readonly dispid 202;
    property BaseImp: Double readonly dispid 203;
    property Alic: Double readonly dispid 204;
    property Importe: Double readonly dispid 205;
  end;

// *********************************************************************//
// Interface: IAlicIva
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {ADE1B3EE-2618-461B-B8D3-F048B400330A}
// *********************************************************************//
  IAlicIva = interface(IDispatch)
    ['{ADE1B3EE-2618-461B-B8D3-F048B400330A}']
    function Get_Id: Integer; safecall;
    function Get_BaseImp: Double; safecall;
    function Get_Importe: Double; safecall;
    property Id: Integer read Get_Id;
    property BaseImp: Double read Get_BaseImp;
    property Importe: Double read Get_Importe;
  end;

// *********************************************************************//
// DispIntf:  IAlicIvaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {ADE1B3EE-2618-461B-B8D3-F048B400330A}
// *********************************************************************//
  IAlicIvaDisp = dispinterface
    ['{ADE1B3EE-2618-461B-B8D3-F048B400330A}']
    property Id: Integer readonly dispid 201;
    property BaseImp: Double readonly dispid 202;
    property Importe: Double readonly dispid 203;
  end;

// *********************************************************************//
// Interface: IOpcional
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7689C644-3F89-44FE-97CF-EAF233A262C8}
// *********************************************************************//
  IOpcional = interface(IDispatch)
    ['{7689C644-3F89-44FE-97CF-EAF233A262C8}']
    function Get_Id: WideString; safecall;
    function Get_Valor: WideString; safecall;
    property Id: WideString read Get_Id;
    property Valor: WideString read Get_Valor;
  end;

// *********************************************************************//
// DispIntf:  IOpcionalDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7689C644-3F89-44FE-97CF-EAF233A262C8}
// *********************************************************************//
  IOpcionalDisp = dispinterface
    ['{7689C644-3F89-44FE-97CF-EAF233A262C8}']
    property Id: WideString readonly dispid 201;
    property Valor: WideString readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IObs
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3417F5A9-B0F6-4CF9-B30B-055E17860895}
// *********************************************************************//
  IObs = interface(IDispatch)
    ['{3417F5A9-B0F6-4CF9-B30B-055E17860895}']
    function Get_Code: Integer; safecall;
    function Get_Msg: WideString; safecall;
    property Code: Integer read Get_Code;
    property Msg: WideString read Get_Msg;
  end;

// *********************************************************************//
// DispIntf:  IObsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3417F5A9-B0F6-4CF9-B30B-055E17860895}
// *********************************************************************//
  IObsDisp = dispinterface
    ['{3417F5A9-B0F6-4CF9-B30B-055E17860895}']
    property Code: Integer readonly dispid 201;
    property Msg: WideString readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IContribuyente
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {19A25CC6-4F15-4C2E-AF88-7AD7901B23A9}
// *********************************************************************//
  IContribuyente = interface(IDispatch)
    ['{19A25CC6-4F15-4C2E-AF88-7AD7901B23A9}']
    function Get_idPersona: WideString; safecall;
    function Get_tipoPersona: WideString; safecall;
    function Get_tipoClave: WideString; safecall;
    function Get_estadoClave: WideString; safecall;
    function Get_nombre: WideString; safecall;
    function Get_tipoDocumento: WideString; safecall;
    function Get_numeroDocumento: WideString; safecall;
    function Get_domicilioFiscal: IDomicilio; safecall;
    function Get_idDependencia: Integer; safecall;
    function Get_mesCierre: Integer; safecall;
    function Get_fechaInscripcion: WideString; safecall;
    function Get_idCatAutonomo: Integer; safecall;
    function Get_impuestosCount: Integer; safecall;
    function impuestos(Indice: Integer): Integer; safecall;
    function categoriasMonotributoCount: Integer; safecall;
    function categoriasMonotributo(Indice: Integer): Integer; safecall;
    function Get_actividadesCount: Integer; safecall;
    function actividades(Indice: Integer): Integer; safecall;
    function Get_condicionIVA: TipoResponsable; safecall;
    function Get_condicionIVADesc: WideString; safecall;
    function Get_SolicitarConstanciaInscripcion: OLE_CANCELBOOL; safecall;
    function actividadesDesc(Inidice: Integer): WideString; safecall;
    function Get_observaciones: WideString; safecall;
    function Get_nombreSimple: WideString; safecall;
    function Get_apellido: WideString; safecall;
    property idPersona: WideString read Get_idPersona;
    property tipoPersona: WideString read Get_tipoPersona;
    property tipoClave: WideString read Get_tipoClave;
    property estadoClave: WideString read Get_estadoClave;
    property nombre: WideString read Get_nombre;
    property tipoDocumento: WideString read Get_tipoDocumento;
    property numeroDocumento: WideString read Get_numeroDocumento;
    property domicilioFiscal: IDomicilio read Get_domicilioFiscal;
    property idDependencia: Integer read Get_idDependencia;
    property mesCierre: Integer read Get_mesCierre;
    property fechaInscripcion: WideString read Get_fechaInscripcion;
    property idCatAutonomo: Integer read Get_idCatAutonomo;
    property impuestosCount: Integer read Get_impuestosCount;
    property actividadesCount: Integer read Get_actividadesCount;
    property condicionIVA: TipoResponsable read Get_condicionIVA;
    property condicionIVADesc: WideString read Get_condicionIVADesc;
    property SolicitarConstanciaInscripcion: OLE_CANCELBOOL read Get_SolicitarConstanciaInscripcion;
    property observaciones: WideString read Get_observaciones;
    property nombreSimple: WideString read Get_nombreSimple;
    property apellido: WideString read Get_apellido;
  end;

// *********************************************************************//
// DispIntf:  IContribuyenteDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {19A25CC6-4F15-4C2E-AF88-7AD7901B23A9}
// *********************************************************************//
  IContribuyenteDisp = dispinterface
    ['{19A25CC6-4F15-4C2E-AF88-7AD7901B23A9}']
    property idPersona: WideString readonly dispid 201;
    property tipoPersona: WideString readonly dispid 202;
    property tipoClave: WideString readonly dispid 203;
    property estadoClave: WideString readonly dispid 204;
    property nombre: WideString readonly dispid 205;
    property tipoDocumento: WideString readonly dispid 206;
    property numeroDocumento: WideString readonly dispid 207;
    property domicilioFiscal: IDomicilio readonly dispid 208;
    property idDependencia: Integer readonly dispid 209;
    property mesCierre: Integer readonly dispid 210;
    property fechaInscripcion: WideString readonly dispid 211;
    property idCatAutonomo: Integer readonly dispid 212;
    property impuestosCount: Integer readonly dispid 213;
    function impuestos(Indice: Integer): Integer; dispid 214;
    function categoriasMonotributoCount: Integer; dispid 215;
    function categoriasMonotributo(Indice: Integer): Integer; dispid 216;
    property actividadesCount: Integer readonly dispid 217;
    function actividades(Indice: Integer): Integer; dispid 218;
    property condicionIVA: TipoResponsable readonly dispid 219;
    property condicionIVADesc: WideString readonly dispid 220;
    property SolicitarConstanciaInscripcion: OLE_CANCELBOOL readonly dispid 221;
    function actividadesDesc(Inidice: Integer): WideString; dispid 222;
    property observaciones: WideString readonly dispid 224;
    property nombreSimple: WideString readonly dispid 223;
    property apellido: WideString readonly dispid 225;
  end;

// *********************************************************************//
// Interface: IDomicilio
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EC378410-896F-4CF2-84A8-53E61AE3D6CF}
// *********************************************************************//
  IDomicilio = interface(IDispatch)
    ['{EC378410-896F-4CF2-84A8-53E61AE3D6CF}']
    function Get_direccion: WideString; safecall;
    function Get_localidad: WideString; safecall;
    function Get_codPostal: WideString; safecall;
    function Get_idProvincia: Integer; safecall;
    function Get_provincia: WideString; safecall;
    property direccion: WideString read Get_direccion;
    property localidad: WideString read Get_localidad;
    property codPostal: WideString read Get_codPostal;
    property idProvincia: Integer read Get_idProvincia;
    property provincia: WideString read Get_provincia;
  end;

// *********************************************************************//
// DispIntf:  IDomicilioDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EC378410-896F-4CF2-84A8-53E61AE3D6CF}
// *********************************************************************//
  IDomicilioDisp = dispinterface
    ['{EC378410-896F-4CF2-84A8-53E61AE3D6CF}']
    property direccion: WideString readonly dispid 201;
    property localidad: WideString readonly dispid 202;
    property codPostal: WideString readonly dispid 203;
    property idProvincia: Integer readonly dispid 204;
    property provincia: WideString readonly dispid 205;
  end;

// *********************************************************************//
// Interface: IwsPadronARBA
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {924DCE98-B918-42E4-A00A-76FD1D8D483A}
// *********************************************************************//
  IwsPadronARBA = interface(IDispatch)
    ['{924DCE98-B918-42E4-A00A-76FD1D8D483A}']
    function ConsultaAlicuota(const fechaDesde: WideString; const fechaHasta: WideString; 
                              CUIT: Double): OLE_CANCELBOOL; safecall;
    function Get_User: WideString; safecall;
    procedure Set_User(const Value: WideString); safecall;
    function Get_Password: WideString; safecall;
    procedure Set_Password(const Value: WideString); safecall;
    function Get_ConsultaAlicuotaRespuesta: IConsultaAlicuotaRespuesta; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_ModoProduccion: OLE_CANCELBOOL; safecall;
    procedure Set_ModoProduccion(Value: OLE_CANCELBOOL); safecall;
    property User: WideString read Get_User write Set_User;
    property Password: WideString read Get_Password write Set_Password;
    property ConsultaAlicuotaRespuesta: IConsultaAlicuotaRespuesta read Get_ConsultaAlicuotaRespuesta;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property ModoProduccion: OLE_CANCELBOOL read Get_ModoProduccion write Set_ModoProduccion;
  end;

// *********************************************************************//
// DispIntf:  IwsPadronARBADisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {924DCE98-B918-42E4-A00A-76FD1D8D483A}
// *********************************************************************//
  IwsPadronARBADisp = dispinterface
    ['{924DCE98-B918-42E4-A00A-76FD1D8D483A}']
    function ConsultaAlicuota(const fechaDesde: WideString; const fechaHasta: WideString; 
                              CUIT: Double): OLE_CANCELBOOL; dispid 201;
    property User: WideString dispid 202;
    property Password: WideString dispid 203;
    property ConsultaAlicuotaRespuesta: IConsultaAlicuotaRespuesta readonly dispid 204;
    property ErrorCode: Integer readonly dispid 205;
    property ErrorDesc: WideString readonly dispid 206;
    property ModoProduccion: OLE_CANCELBOOL dispid 207;
  end;

// *********************************************************************//
// Interface: IConsultaAlicuotaRespuesta
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2589E4FF-0788-4FEF-9565-0F05095F1356}
// *********************************************************************//
  IConsultaAlicuotaRespuesta = interface(IDispatch)
    ['{2589E4FF-0788-4FEF-9565-0F05095F1356}']
    function Get_AlicuotaPercepcion: Double; safecall;
    function Get_AlicuotaRetencion: Double; safecall;
    function Get_GrupoPercepcion: Integer; safecall;
    function Get_GrupoRetencion: Integer; safecall;
    property AlicuotaPercepcion: Double read Get_AlicuotaPercepcion;
    property AlicuotaRetencion: Double read Get_AlicuotaRetencion;
    property GrupoPercepcion: Integer read Get_GrupoPercepcion;
    property GrupoRetencion: Integer read Get_GrupoRetencion;
  end;

// *********************************************************************//
// DispIntf:  IConsultaAlicuotaRespuestaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2589E4FF-0788-4FEF-9565-0F05095F1356}
// *********************************************************************//
  IConsultaAlicuotaRespuestaDisp = dispinterface
    ['{2589E4FF-0788-4FEF-9565-0F05095F1356}']
    property AlicuotaPercepcion: Double readonly dispid 201;
    property AlicuotaRetencion: Double readonly dispid 202;
    property GrupoPercepcion: Integer readonly dispid 203;
    property GrupoRetencion: Integer readonly dispid 204;
  end;

// *********************************************************************//
// Interface: ICertificado
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {CAD1F637-CD57-45DF-8A39-EB2227E34D93}
// *********************************************************************//
  ICertificado = interface(IDispatch)
    ['{CAD1F637-CD57-45DF-8A39-EB2227E34D93}']
    function CargarInformacionCertificado(const ArchivoCertificado: WideString; 
                                          const ArchivoClavePrivada: WideString): OLE_CANCELBOOL; safecall;
    function GenerarNuevoCertificado(const O: WideString; const CN: WideString; CUIT: Double; 
                                     const ArchivoSolicitud: WideString; 
                                     const ArchivoClavePrivada: WideString): OLE_CANCELBOOL; safecall;
    function RenovarCertificado(const ArchivoSolicitud: WideString): OLE_CANCELBOOL; safecall;
    procedure MostrarInformacionCertificado; safecall;
    procedure MostrarGenerarCertificado; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_IC_Organizacion: WideString; safecall;
    function Get_IC_NombreComun: WideString; safecall;
    function Get_IC_FechaVencimiento: WideString; safecall;
    function Get_IC_CUIT: Double; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property IC_Organizacion: WideString read Get_IC_Organizacion;
    property IC_NombreComun: WideString read Get_IC_NombreComun;
    property IC_FechaVencimiento: WideString read Get_IC_FechaVencimiento;
    property IC_CUIT: Double read Get_IC_CUIT;
  end;

// *********************************************************************//
// DispIntf:  ICertificadoDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {CAD1F637-CD57-45DF-8A39-EB2227E34D93}
// *********************************************************************//
  ICertificadoDisp = dispinterface
    ['{CAD1F637-CD57-45DF-8A39-EB2227E34D93}']
    function CargarInformacionCertificado(const ArchivoCertificado: WideString; 
                                          const ArchivoClavePrivada: WideString): OLE_CANCELBOOL; dispid 201;
    function GenerarNuevoCertificado(const O: WideString; const CN: WideString; CUIT: Double; 
                                     const ArchivoSolicitud: WideString; 
                                     const ArchivoClavePrivada: WideString): OLE_CANCELBOOL; dispid 203;
    function RenovarCertificado(const ArchivoSolicitud: WideString): OLE_CANCELBOOL; dispid 206;
    procedure MostrarInformacionCertificado; dispid 210;
    procedure MostrarGenerarCertificado; dispid 211;
    property ErrorCode: Integer readonly dispid 204;
    property ErrorDesc: WideString readonly dispid 205;
    property IC_Organizacion: WideString readonly dispid 202;
    property IC_NombreComun: WideString readonly dispid 207;
    property IC_FechaVencimiento: WideString readonly dispid 209;
    property IC_CUIT: Double readonly dispid 208;
  end;

// *********************************************************************//
// Interface: Iwscdc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {201C6546-D660-4171-A3D3-839583F7969E}
// *********************************************************************//
  Iwscdc = interface(IDispatch)
    ['{201C6546-D660-4171-A3D3-839583F7969E}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; safecall;
    function ComprobanteConstatar(const CbteModo: WideString; CuitEmisor: Double; PtoVta: Integer; 
                                  CbteTipo: Integer; CbteNro: Double; const CbteFch: WideString; 
                                  Imptotal: Double; const CodAutorizacion: WideString; 
                                  const DocTipoReceptor: WideString; 
                                  const DocNroReceptor: WideString): OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_URL: WideString; safecall;
    procedure Set_URL(const Value: WideString); safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property URL: WideString read Get_URL write Set_URL;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
  end;

// *********************************************************************//
// DispIntf:  IwscdcDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {201C6546-D660-4171-A3D3-839583F7969E}
// *********************************************************************//
  IwscdcDisp = dispinterface
    ['{201C6546-D660-4171-A3D3-839583F7969E}']
    function login(const Certificado: WideString; const ClavePrivada: WideString; 
                   const URL: WideString): OLE_CANCELBOOL; dispid 201;
    function ComprobanteConstatar(const CbteModo: WideString; CuitEmisor: Double; PtoVta: Integer; 
                                  CbteTipo: Integer; CbteNro: Double; const CbteFch: WideString; 
                                  Imptotal: Double; const CodAutorizacion: WideString; 
                                  const DocTipoReceptor: WideString; 
                                  const DocNroReceptor: WideString): OLE_CANCELBOOL; dispid 202;
    property ErrorCode: Integer readonly dispid 203;
    property ErrorDesc: WideString readonly dispid 204;
    property URL: WideString dispid 205;
    property CUIT: Double dispid 206;
    property Depurar: OLE_CANCELBOOL dispid 207;
  end;

// *********************************************************************//
// Interface: IBarcode
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {01F6CFB9-A47D-401E-8A89-1C3962BB9364}
// *********************************************************************//
  IBarcode = interface(IDispatch)
    ['{01F6CFB9-A47D-401E-8A89-1C3962BB9364}']
    procedure GenerarCodigo(CUIT: Double; TipoCbte: Integer; PtoVta: Integer; 
                            const Cae: WideString; const Vto: WideString; 
                            const ArchivoDestino: WideString); safecall;
    function Get_Modulo: Integer; safecall;
    procedure Set_Modulo(Value: Integer); safecall;
    function Get_Proporcion: Double; safecall;
    procedure Set_Proporcion(Value: Double); safecall;
    function Get_Altura: Integer; safecall;
    procedure Set_Altura(Value: Integer); safecall;
    function Get_MostrarTexto: OLE_CANCELBOOL; safecall;
    procedure Set_MostrarTexto(Value: OLE_CANCELBOOL); safecall;
    function Get_TamanioFuente: Integer; safecall;
    procedure Set_TamanioFuente(Value: Integer); safecall;
    function Get_Texto: WideString; safecall;
    function Interleave25(const Texto: WideString; const ArchivoDestino: WideString): OLE_CANCELBOOL; safecall;
    property Modulo: Integer read Get_Modulo write Set_Modulo;
    property Proporcion: Double read Get_Proporcion write Set_Proporcion;
    property Altura: Integer read Get_Altura write Set_Altura;
    property MostrarTexto: OLE_CANCELBOOL read Get_MostrarTexto write Set_MostrarTexto;
    property TamanioFuente: Integer read Get_TamanioFuente write Set_TamanioFuente;
    property Texto: WideString read Get_Texto;
  end;

// *********************************************************************//
// DispIntf:  IBarcodeDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {01F6CFB9-A47D-401E-8A89-1C3962BB9364}
// *********************************************************************//
  IBarcodeDisp = dispinterface
    ['{01F6CFB9-A47D-401E-8A89-1C3962BB9364}']
    procedure GenerarCodigo(CUIT: Double; TipoCbte: Integer; PtoVta: Integer; 
                            const Cae: WideString; const Vto: WideString; 
                            const ArchivoDestino: WideString); dispid 201;
    property Modulo: Integer dispid 202;
    property Proporcion: Double dispid 203;
    property Altura: Integer dispid 204;
    property MostrarTexto: OLE_CANCELBOOL dispid 205;
    property TamanioFuente: Integer dispid 206;
    property Texto: WideString readonly dispid 207;
    function Interleave25(const Texto: WideString; const ArchivoDestino: WideString): OLE_CANCELBOOL; dispid 208;
  end;

// *********************************************************************//
// Interface: Iwsct
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {161A74B4-F8B8-408F-934B-2D2D32E492E2}
// *********************************************************************//
  Iwsct = interface(IDispatch)
    ['{161A74B4-F8B8-408F-934B-2D2D32E492E2}']
    procedure AgregaFactura(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                            numeroComprobante: Double; const fechaEmision: WideString; 
                            const codigoTipoAutorizacion: WideString; codigoAutorizacion: Double; 
                            const fechaVencimiento: WideString; codigoTipoDocumento: Integer; 
                            const numeroDocumento: WideString; const idImpositivo: WideString; 
                            codigoPais: Integer; const domicilioReceptor: WideString; 
                            codigoRelacionEmisorReceptor: Integer; importeGravado: Double; 
                            importeNoGravado: Double; importeExento: Double; 
                            importeOtrosTributos: Double; importeReintegro: Double; 
                            importeTotal: Double; const codigoMoneda: WideString; 
                            cotizacionMoneda: Double; const observaciones: WideString); safecall;
    procedure AgregaItem(Tipo: Integer; codigoTurismo: Integer; const codigo: WideString; 
                         const descripcion: WideString; codigoAlicuotaIVA: Integer; 
                         importeIVA: Double; importeItem: Double); safecall;
    procedure AgregaComprobanteAsociado(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                                        numeroComprobante: Double); safecall;
    procedure AgregaTributo(codigo: Integer; const descripcion: WideString; baseImponible: Double; 
                            Importe: Double); safecall;
    procedure AgregaIVA(codigo: Integer; Importe: Double); safecall;
    procedure AgregaDatoAdicional(T: Integer; const C1: WideString; const C2: WideString; 
                                  const C3: WideString; const C4: WideString; const C5: WideString; 
                                  const C6: WideString); safecall;
    procedure AgregaFormaDePago(codigo: Integer; tipoTarjeta: Integer; numeroTarjeta: Double; 
                                const swiftCode: WideString; tipoCuenta: Integer; 
                                numeroCuenta: Double); safecall;
    procedure Reset; safecall;
    function Autorizar: OLE_CANCELBOOL; safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    function ConsultarUltimoComprobante(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer): OLE_CANCELBOOL; safecall;
    function Get_ConsultarUltimoComprobanteNumero: Integer; safecall;
    function Get_ConsultarUltimoComprobanteFecha: WideString; safecall;
    function DescargarCodigos(const NombreArchivo: WideString): OLE_CANCELBOOL; safecall;
    function Get_AutorizarCAE: Double; safecall;
    function Get_AutorizarVencimiento: WideString; safecall;
    function Get_AutorizarObservaciones: WideString; safecall;
    function Get_ModoProduccion: OLE_CANCELBOOL; safecall;
    procedure Set_ModoProduccion(Value: OLE_CANCELBOOL); safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property ConsultarUltimoComprobanteNumero: Integer read Get_ConsultarUltimoComprobanteNumero;
    property ConsultarUltimoComprobanteFecha: WideString read Get_ConsultarUltimoComprobanteFecha;
    property AutorizarCAE: Double read Get_AutorizarCAE;
    property AutorizarVencimiento: WideString read Get_AutorizarVencimiento;
    property AutorizarObservaciones: WideString read Get_AutorizarObservaciones;
    property ModoProduccion: OLE_CANCELBOOL read Get_ModoProduccion write Set_ModoProduccion;
  end;

// *********************************************************************//
// DispIntf:  IwsctDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {161A74B4-F8B8-408F-934B-2D2D32E492E2}
// *********************************************************************//
  IwsctDisp = dispinterface
    ['{161A74B4-F8B8-408F-934B-2D2D32E492E2}']
    procedure AgregaFactura(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                            numeroComprobante: Double; const fechaEmision: WideString; 
                            const codigoTipoAutorizacion: WideString; codigoAutorizacion: Double; 
                            const fechaVencimiento: WideString; codigoTipoDocumento: Integer; 
                            const numeroDocumento: WideString; const idImpositivo: WideString; 
                            codigoPais: Integer; const domicilioReceptor: WideString; 
                            codigoRelacionEmisorReceptor: Integer; importeGravado: Double; 
                            importeNoGravado: Double; importeExento: Double; 
                            importeOtrosTributos: Double; importeReintegro: Double; 
                            importeTotal: Double; const codigoMoneda: WideString; 
                            cotizacionMoneda: Double; const observaciones: WideString); dispid 201;
    procedure AgregaItem(Tipo: Integer; codigoTurismo: Integer; const codigo: WideString; 
                         const descripcion: WideString; codigoAlicuotaIVA: Integer; 
                         importeIVA: Double; importeItem: Double); dispid 202;
    procedure AgregaComprobanteAsociado(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer; 
                                        numeroComprobante: Double); dispid 203;
    procedure AgregaTributo(codigo: Integer; const descripcion: WideString; baseImponible: Double; 
                            Importe: Double); dispid 204;
    procedure AgregaIVA(codigo: Integer; Importe: Double); dispid 205;
    procedure AgregaDatoAdicional(T: Integer; const C1: WideString; const C2: WideString; 
                                  const C3: WideString; const C4: WideString; const C5: WideString; 
                                  const C6: WideString); dispid 206;
    procedure AgregaFormaDePago(codigo: Integer; tipoTarjeta: Integer; numeroTarjeta: Double; 
                                const swiftCode: WideString; tipoCuenta: Integer; 
                                numeroCuenta: Double); dispid 207;
    procedure Reset; dispid 208;
    function Autorizar: OLE_CANCELBOOL; dispid 209;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; dispid 210;
    property ErrorCode: Integer readonly dispid 211;
    property ErrorDesc: WideString readonly dispid 212;
    property CUIT: Double dispid 213;
    property Depurar: OLE_CANCELBOOL dispid 214;
    function ConsultarUltimoComprobante(codigoTipoComprobante: Integer; numeroPuntoVenta: Integer): OLE_CANCELBOOL; dispid 215;
    property ConsultarUltimoComprobanteNumero: Integer readonly dispid 216;
    property ConsultarUltimoComprobanteFecha: WideString readonly dispid 217;
    function DescargarCodigos(const NombreArchivo: WideString): OLE_CANCELBOOL; dispid 218;
    property AutorizarCAE: Double readonly dispid 219;
    property AutorizarVencimiento: WideString readonly dispid 220;
    property AutorizarObservaciones: WideString readonly dispid 221;
    property ModoProduccion: OLE_CANCELBOOL dispid 222;
  end;

// *********************************************************************//
// Interface: Iwsfecred
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {32EF8E70-4CB3-40FD-A66C-BBB03E147C37}
// *********************************************************************//
  Iwsfecred = interface(IDispatch)
    ['{32EF8E70-4CB3-40FD-A66C-BBB03E147C37}']
    procedure Dummy; safecall;
    function consultarComprobantes(const rolCUITRepresentada: WideString; CUITContraparte: Double; 
                                   codTipoCmp: Integer; const estadoCmp: WideString; 
                                   const fecha_tipo: WideString; const fecha_desde: WideString; 
                                   const fecha_hasta: WideString; codCtaCte: Double; 
                                   const estadoCtaCte: WideString): OLE_CANCELBOOL; safecall;
    procedure rechazarNotaDC; safecall;
    function consultarCtasCtes(const rolCUITRepresentada: WideString; CUITContraparte: Double; 
                               const fecha: WideString; const estadoCtaCte: WideString): OLE_CANCELBOOL; safecall;
    function consultarCtaCte(codCtaCte: Integer; CuitEmisor: Double; codTipoCmp: Integer; 
                             PtoVta: Integer; nroCmp: Double): OLE_CANCELBOOL; safecall;
    procedure informarCancelacionTotalFECred; safecall;
    function aceptarFECred(const Request: IAceptarFECredRequestTy): OLE_CANCELBOOL; safecall;
    function rechazarFECred(const Request: IRechazarFECredRequestTy): OLE_CANCELBOOL; safecall;
    function informarFacturaAgtDptoCltv(const Request: IInformarFacturaAgtDptoCltvRequestTy): OLE_CANCELBOOL; safecall;
    procedure consultarFacturasAgtDptoCltv; safecall;
    procedure consultarCuentasComitente; safecall;
    function consultarObligadoRecepcion(cuitConsultada: Double): OLE_CANCELBOOL; safecall;
    procedure consultarTiposRetenciones; safecall;
    procedure consultarTiposMotivosRechazo; safecall;
    procedure consultarTiposFormasCancelacion; safecall;
    procedure obtenerRemitos; safecall;
    procedure consultarHistorialEstadosComprobante; safecall;
    procedure consultarHistorialEstadosCtaCte; safecall;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    function Get_Token: WideString; safecall;
    function Get_Sign: WideString; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_CUIT: Double; safecall;
    procedure Set_CUIT(Value: Double); safecall;
    function Get_XMLRequest: WideString; safecall;
    function Get_XMLResponse: WideString; safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    function nuevoAceptarFECredRequestTy: IAceptarFECredRequestTy; safecall;
    function Get_ModoProduccion: OLE_CANCELBOOL; safecall;
    procedure Set_ModoProduccion(Value: OLE_CANCELBOOL); safecall;
    function Get_consultarCmpReturn: IConsultarCmpReturnTy; safecall;
    function nuevoInformarFacturaAgtDptoCltvRequestTy: IInformarFacturaAgtDptoCltvRequestTy; safecall;
    function nuevoRechazarFECredRequestTy: IRechazarFECredRequestTy; safecall;
    function Get_consultarObligadoRecepcionReturn: IconsultarObligadoRecepcionReturnTy; safecall;
    function consultarMontoObligadoRecepcion(cuitConsultada: Double; const fechaEmision: WideString): OLE_CANCELBOOL; safecall;
    function Get_consultarMontoObligadoRecepcionReturn: IConsultarMontoObligadoRecepcionReturnTy; safecall;
    function Get_consultarCtasCtesReturn: IConsultarCtasCtesReturnTy; safecall;
    function Get_consultarCtaCteReturn: IConsultarCtaCteReturnTy; safecall;
    property Token: WideString read Get_Token;
    property Sign: WideString read Get_Sign;
    property ErrorCode: Integer read Get_ErrorCode;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property CUIT: Double read Get_CUIT write Set_CUIT;
    property XMLRequest: WideString read Get_XMLRequest;
    property XMLResponse: WideString read Get_XMLResponse;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property ModoProduccion: OLE_CANCELBOOL read Get_ModoProduccion write Set_ModoProduccion;
    property consultarCmpReturn: IConsultarCmpReturnTy read Get_consultarCmpReturn;
    property consultarObligadoRecepcionReturn: IconsultarObligadoRecepcionReturnTy read Get_consultarObligadoRecepcionReturn;
    property consultarMontoObligadoRecepcionReturn: IConsultarMontoObligadoRecepcionReturnTy read Get_consultarMontoObligadoRecepcionReturn;
    property consultarCtasCtesReturn: IConsultarCtasCtesReturnTy read Get_consultarCtasCtesReturn;
    property consultarCtaCteReturn: IConsultarCtaCteReturnTy read Get_consultarCtaCteReturn;
  end;

// *********************************************************************//
// DispIntf:  IwsfecredDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {32EF8E70-4CB3-40FD-A66C-BBB03E147C37}
// *********************************************************************//
  IwsfecredDisp = dispinterface
    ['{32EF8E70-4CB3-40FD-A66C-BBB03E147C37}']
    procedure Dummy; dispid 201;
    function consultarComprobantes(const rolCUITRepresentada: WideString; CUITContraparte: Double; 
                                   codTipoCmp: Integer; const estadoCmp: WideString; 
                                   const fecha_tipo: WideString; const fecha_desde: WideString; 
                                   const fecha_hasta: WideString; codCtaCte: Double; 
                                   const estadoCtaCte: WideString): OLE_CANCELBOOL; dispid 202;
    procedure rechazarNotaDC; dispid 203;
    function consultarCtasCtes(const rolCUITRepresentada: WideString; CUITContraparte: Double; 
                               const fecha: WideString; const estadoCtaCte: WideString): OLE_CANCELBOOL; dispid 204;
    function consultarCtaCte(codCtaCte: Integer; CuitEmisor: Double; codTipoCmp: Integer; 
                             PtoVta: Integer; nroCmp: Double): OLE_CANCELBOOL; dispid 205;
    procedure informarCancelacionTotalFECred; dispid 206;
    function aceptarFECred(const Request: IAceptarFECredRequestTy): OLE_CANCELBOOL; dispid 207;
    function rechazarFECred(const Request: IRechazarFECredRequestTy): OLE_CANCELBOOL; dispid 208;
    function informarFacturaAgtDptoCltv(const Request: IInformarFacturaAgtDptoCltvRequestTy): OLE_CANCELBOOL; dispid 209;
    procedure consultarFacturasAgtDptoCltv; dispid 210;
    procedure consultarCuentasComitente; dispid 211;
    function consultarObligadoRecepcion(cuitConsultada: Double): OLE_CANCELBOOL; dispid 212;
    procedure consultarTiposRetenciones; dispid 213;
    procedure consultarTiposMotivosRechazo; dispid 214;
    procedure consultarTiposFormasCancelacion; dispid 215;
    procedure obtenerRemitos; dispid 216;
    procedure consultarHistorialEstadosComprobante; dispid 217;
    procedure consultarHistorialEstadosCtaCte; dispid 218;
    function login(const Certificado: WideString; const ClavePrivada: WideString): OLE_CANCELBOOL; dispid 219;
    procedure CargarLicencia(const Licencia: WideString); dispid 220;
    property Token: WideString readonly dispid 221;
    property Sign: WideString readonly dispid 222;
    property ErrorCode: Integer readonly dispid 223;
    property ErrorDesc: WideString readonly dispid 224;
    property CUIT: Double dispid 225;
    property XMLRequest: WideString readonly dispid 226;
    property XMLResponse: WideString readonly dispid 227;
    property Depurar: OLE_CANCELBOOL dispid 228;
    function nuevoAceptarFECredRequestTy: IAceptarFECredRequestTy; dispid 229;
    property ModoProduccion: OLE_CANCELBOOL dispid 230;
    property consultarCmpReturn: IConsultarCmpReturnTy readonly dispid 231;
    function nuevoInformarFacturaAgtDptoCltvRequestTy: IInformarFacturaAgtDptoCltvRequestTy; dispid 232;
    function nuevoRechazarFECredRequestTy: IRechazarFECredRequestTy; dispid 233;
    property consultarObligadoRecepcionReturn: IconsultarObligadoRecepcionReturnTy readonly dispid 234;
    function consultarMontoObligadoRecepcion(cuitConsultada: Double; const fechaEmision: WideString): OLE_CANCELBOOL; dispid 235;
    property consultarMontoObligadoRecepcionReturn: IConsultarMontoObligadoRecepcionReturnTy readonly dispid 236;
    property consultarCtasCtesReturn: IConsultarCtasCtesReturnTy readonly dispid 237;
    property consultarCtaCteReturn: IConsultarCtaCteReturnTy readonly dispid 238;
  end;

// *********************************************************************//
// Interface: IIdCtaCteTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C9194512-99E1-4404-85AB-6218E498CEED}
// *********************************************************************//
  IIdCtaCteTy = interface(IDispatch)
    ['{C9194512-99E1-4404-85AB-6218E498CEED}']
    function Get_codCtaCte: Double; safecall;
    procedure Set_codCtaCte(Value: Double); safecall;
    function Get_idFactura: IdComprobanteTy; safecall;
    procedure Set_idFactura(const Value: IdComprobanteTy); safecall;
    property codCtaCte: Double read Get_codCtaCte write Set_codCtaCte;
    property idFactura: IdComprobanteTy read Get_idFactura write Set_idFactura;
  end;

// *********************************************************************//
// DispIntf:  IIdCtaCteTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C9194512-99E1-4404-85AB-6218E498CEED}
// *********************************************************************//
  IIdCtaCteTyDisp = dispinterface
    ['{C9194512-99E1-4404-85AB-6218E498CEED}']
    property codCtaCte: Double dispid 201;
    property idFactura: IdComprobanteTy dispid 202;
  end;

// *********************************************************************//
// Interface: IIdComprobanteTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3ABD3582-6764-4A05-BFDE-CFED3D4A1143}
// *********************************************************************//
  IIdComprobanteTy = interface(IDispatch)
    ['{3ABD3582-6764-4A05-BFDE-CFED3D4A1143}']
    function Get_CuitEmisor: Double; safecall;
    procedure Set_CuitEmisor(Value: Double); safecall;
    function Get_codTipoCmp: Integer; safecall;
    procedure Set_codTipoCmp(Value: Integer); safecall;
    function Get_PtoVta: Integer; safecall;
    procedure Set_PtoVta(Value: Integer); safecall;
    function Get_nroCmp: Double; safecall;
    procedure Set_nroCmp(Value: Double); safecall;
    property CuitEmisor: Double read Get_CuitEmisor write Set_CuitEmisor;
    property codTipoCmp: Integer read Get_codTipoCmp write Set_codTipoCmp;
    property PtoVta: Integer read Get_PtoVta write Set_PtoVta;
    property nroCmp: Double read Get_nroCmp write Set_nroCmp;
  end;

// *********************************************************************//
// DispIntf:  IIdComprobanteTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3ABD3582-6764-4A05-BFDE-CFED3D4A1143}
// *********************************************************************//
  IIdComprobanteTyDisp = dispinterface
    ['{3ABD3582-6764-4A05-BFDE-CFED3D4A1143}']
    property CuitEmisor: Double dispid 201;
    property codTipoCmp: Integer dispid 202;
    property PtoVta: Integer dispid 203;
    property nroCmp: Double dispid 204;
  end;

// *********************************************************************//
// Interface: IAceptarFECredRequestTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F0324362-5DE0-4A53-B253-D18C37D5FD5C}
// *********************************************************************//
  IAceptarFECredRequestTy = interface(IDispatch)
    ['{F0324362-5DE0-4A53-B253-D18C37D5FD5C}']
    procedure idCtaCte(codCtaCte: Double); safecall;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); safecall;
    procedure arrayConfirmarNotasDC(acepta: OLE_CANCELBOOL; CuitEmisor: Double; 
                                    codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); safecall;
    procedure arrayFormasCancelacion(codigo: Integer; const descripcion: WideString); safecall;
    procedure arrayRetenciones(codTipo: Integer; Importe: Double; Porcentaje: Double; 
                               const descMotivo: WideString); safecall;
    procedure arrayAjustesOperacion(codigo: Integer; Importe: Double); safecall;
    function Get_tipoCancelacion: WideString; safecall;
    procedure Set_tipoCancelacion(const Value: WideString); safecall;
    function Get_importeCancelado: Double; safecall;
    procedure Set_importeCancelado(Value: Double); safecall;
    function Get_importeTotalRetPesos: Double; safecall;
    procedure Set_importeTotalRetPesos(Value: Double); safecall;
    function Get_importeEmbargoPesos: Double; safecall;
    procedure Set_importeEmbargoPesos(Value: Double); safecall;
    function Get_saldoAceptado: Double; safecall;
    procedure Set_saldoAceptado(Value: Double); safecall;
    function Get_codMoneda: WideString; safecall;
    procedure Set_codMoneda(const Value: WideString); safecall;
    function Get_cotizacionMonedaUlt: Double; safecall;
    procedure Set_cotizacionMonedaUlt(Value: Double); safecall;
    property tipoCancelacion: WideString read Get_tipoCancelacion write Set_tipoCancelacion;
    property importeCancelado: Double read Get_importeCancelado write Set_importeCancelado;
    property importeTotalRetPesos: Double read Get_importeTotalRetPesos write Set_importeTotalRetPesos;
    property importeEmbargoPesos: Double read Get_importeEmbargoPesos write Set_importeEmbargoPesos;
    property saldoAceptado: Double read Get_saldoAceptado write Set_saldoAceptado;
    property codMoneda: WideString read Get_codMoneda write Set_codMoneda;
    property cotizacionMonedaUlt: Double read Get_cotizacionMonedaUlt write Set_cotizacionMonedaUlt;
  end;

// *********************************************************************//
// DispIntf:  IAceptarFECredRequestTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F0324362-5DE0-4A53-B253-D18C37D5FD5C}
// *********************************************************************//
  IAceptarFECredRequestTyDisp = dispinterface
    ['{F0324362-5DE0-4A53-B253-D18C37D5FD5C}']
    procedure idCtaCte(codCtaCte: Double); dispid 201;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); dispid 213;
    procedure arrayConfirmarNotasDC(acepta: OLE_CANCELBOOL; CuitEmisor: Double; 
                                    codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); dispid 202;
    procedure arrayFormasCancelacion(codigo: Integer; const descripcion: WideString); dispid 203;
    procedure arrayRetenciones(codTipo: Integer; Importe: Double; Porcentaje: Double; 
                               const descMotivo: WideString); dispid 204;
    procedure arrayAjustesOperacion(codigo: Integer; Importe: Double); dispid 205;
    property tipoCancelacion: WideString dispid 206;
    property importeCancelado: Double dispid 207;
    property importeTotalRetPesos: Double dispid 208;
    property importeEmbargoPesos: Double dispid 209;
    property saldoAceptado: Double dispid 210;
    property codMoneda: WideString dispid 211;
    property cotizacionMonedaUlt: Double dispid 212;
  end;

// *********************************************************************//
// Interface: IConsultarCmpReturnTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BB315EBC-4D6F-4542-BC1D-B4878E91A9EF}
// *********************************************************************//
  IConsultarCmpReturnTy = interface(IDispatch)
    ['{BB315EBC-4D6F-4542-BC1D-B4878E91A9EF}']
    function Get_arrayComprobantes(Indice: Integer): IComprobanteTy; safecall;
    function Get_arrayComprobantesCount: Integer; safecall;
    property arrayComprobantes[Indice: Integer]: IComprobanteTy read Get_arrayComprobantes;
    property arrayComprobantesCount: Integer read Get_arrayComprobantesCount;
  end;

// *********************************************************************//
// DispIntf:  IConsultarCmpReturnTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BB315EBC-4D6F-4542-BC1D-B4878E91A9EF}
// *********************************************************************//
  IConsultarCmpReturnTyDisp = dispinterface
    ['{BB315EBC-4D6F-4542-BC1D-B4878E91A9EF}']
    property arrayComprobantes[Indice: Integer]: IComprobanteTy readonly dispid 201;
    property arrayComprobantesCount: Integer readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IComprobanteTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E4EF8C5E-D0D5-4C0A-A1AA-AFDC93A8D4E7}
// *********************************************************************//
  IComprobanteTy = interface(IDispatch)
    ['{E4EF8C5E-D0D5-4C0A-A1AA-AFDC93A8D4E7}']
    function Get_CuitEmisor: Double; safecall;
    function Get_razonSocialEmi: WideString; safecall;
    function Get_codTipoCmp: Integer; safecall;
    function Get_PtoVta: Integer; safecall;
    function Get_nroCmp: Double; safecall;
    function Get_cuitReceptor: Double; safecall;
    function Get_razonSocialRecep: WideString; safecall;
    function Get_tipoCodAuto: WideString; safecall;
    function Get_CodAutorizacion: Double; safecall;
    function Get_fechaEmision: WideString; safecall;
    function Get_fechaPuestaDispo: WideString; safecall;
    function Get_fechaVenPago: WideString; safecall;
    function Get_fechaVenAcep: WideString; safecall;
    function Get_importeTotal: Double; safecall;
    function Get_codMoneda: WideString; safecall;
    function Get_cotizacionMoneda: Double; safecall;
    function Get_CBUEmisor: WideString; safecall;
    function Get_AliasEmisor: WideString; safecall;
    function Get_esAnulacion: OLE_CANCELBOOL; safecall;
    function Get_esPostAceptacion: OLE_CANCELBOOL; safecall;
    function Get_idComprobanteAsociado: IIdComprobanteTy; safecall;
    function Get_referenciasComerciales: WideString; safecall;
    function Get_arraySubtotalesIVA(Indice: Integer): ISubtotalIVATy; safecall;
    function Get_arraySubtotalesIVACount: Integer; safecall;
    function Get_arrayOtrosTributos: IOtroTributoTy; safecall;
    function Get_arrayOtrosTributosCount: Integer; safecall;
    function Get_arrayItems: IItemTy; safecall;
    function Get_arrayItemsCount: Integer; safecall;
    function Get_datosGenerales: WideString; safecall;
    function Get_datosComerciales: WideString; safecall;
    function Get_leyendaComercial: WideString; safecall;
    function Get_codCtaCte: Double; safecall;
    function Get_estado_estado: WideString; safecall;
    function Get_estado_fecha: WideString; safecall;
    function Get_tipoAcep: WideString; safecall;
    function Get_fechaHoraAcep: WideString; safecall;
    function Get_arrayMotivosRechazo: IMotivoRechazoTy; safecall;
    function Get_arrayMotivosRechazoCount: Integer; safecall;
    function Get_infoAgDtpoCltv: OLE_CANCELBOOL; safecall;
    function Get_fechaInfoAgDptoCltv: WideString; safecall;
    function Get_idPagoAgDptoCltv: WideString; safecall;
    function Get_CBUdePago: WideString; safecall;
    property CuitEmisor: Double read Get_CuitEmisor;
    property razonSocialEmi: WideString read Get_razonSocialEmi;
    property codTipoCmp: Integer read Get_codTipoCmp;
    property PtoVta: Integer read Get_PtoVta;
    property nroCmp: Double read Get_nroCmp;
    property cuitReceptor: Double read Get_cuitReceptor;
    property razonSocialRecep: WideString read Get_razonSocialRecep;
    property tipoCodAuto: WideString read Get_tipoCodAuto;
    property CodAutorizacion: Double read Get_CodAutorizacion;
    property fechaEmision: WideString read Get_fechaEmision;
    property fechaPuestaDispo: WideString read Get_fechaPuestaDispo;
    property fechaVenPago: WideString read Get_fechaVenPago;
    property fechaVenAcep: WideString read Get_fechaVenAcep;
    property importeTotal: Double read Get_importeTotal;
    property codMoneda: WideString read Get_codMoneda;
    property cotizacionMoneda: Double read Get_cotizacionMoneda;
    property CBUEmisor: WideString read Get_CBUEmisor;
    property AliasEmisor: WideString read Get_AliasEmisor;
    property esAnulacion: OLE_CANCELBOOL read Get_esAnulacion;
    property esPostAceptacion: OLE_CANCELBOOL read Get_esPostAceptacion;
    property idComprobanteAsociado: IIdComprobanteTy read Get_idComprobanteAsociado;
    property referenciasComerciales: WideString read Get_referenciasComerciales;
    property arraySubtotalesIVA[Indice: Integer]: ISubtotalIVATy read Get_arraySubtotalesIVA;
    property arraySubtotalesIVACount: Integer read Get_arraySubtotalesIVACount;
    property arrayOtrosTributos: IOtroTributoTy read Get_arrayOtrosTributos;
    property arrayOtrosTributosCount: Integer read Get_arrayOtrosTributosCount;
    property arrayItems: IItemTy read Get_arrayItems;
    property arrayItemsCount: Integer read Get_arrayItemsCount;
    property datosGenerales: WideString read Get_datosGenerales;
    property datosComerciales: WideString read Get_datosComerciales;
    property leyendaComercial: WideString read Get_leyendaComercial;
    property codCtaCte: Double read Get_codCtaCte;
    property estado_estado: WideString read Get_estado_estado;
    property estado_fecha: WideString read Get_estado_fecha;
    property tipoAcep: WideString read Get_tipoAcep;
    property fechaHoraAcep: WideString read Get_fechaHoraAcep;
    property arrayMotivosRechazo: IMotivoRechazoTy read Get_arrayMotivosRechazo;
    property arrayMotivosRechazoCount: Integer read Get_arrayMotivosRechazoCount;
    property infoAgDtpoCltv: OLE_CANCELBOOL read Get_infoAgDtpoCltv;
    property fechaInfoAgDptoCltv: WideString read Get_fechaInfoAgDptoCltv;
    property idPagoAgDptoCltv: WideString read Get_idPagoAgDptoCltv;
    property CBUdePago: WideString read Get_CBUdePago;
  end;

// *********************************************************************//
// DispIntf:  IComprobanteTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E4EF8C5E-D0D5-4C0A-A1AA-AFDC93A8D4E7}
// *********************************************************************//
  IComprobanteTyDisp = dispinterface
    ['{E4EF8C5E-D0D5-4C0A-A1AA-AFDC93A8D4E7}']
    property CuitEmisor: Double readonly dispid 201;
    property razonSocialEmi: WideString readonly dispid 202;
    property codTipoCmp: Integer readonly dispid 203;
    property PtoVta: Integer readonly dispid 204;
    property nroCmp: Double readonly dispid 205;
    property cuitReceptor: Double readonly dispid 206;
    property razonSocialRecep: WideString readonly dispid 207;
    property tipoCodAuto: WideString readonly dispid 208;
    property CodAutorizacion: Double readonly dispid 209;
    property fechaEmision: WideString readonly dispid 210;
    property fechaPuestaDispo: WideString readonly dispid 211;
    property fechaVenPago: WideString readonly dispid 212;
    property fechaVenAcep: WideString readonly dispid 213;
    property importeTotal: Double readonly dispid 214;
    property codMoneda: WideString readonly dispid 215;
    property cotizacionMoneda: Double readonly dispid 216;
    property CBUEmisor: WideString readonly dispid 217;
    property AliasEmisor: WideString readonly dispid 218;
    property esAnulacion: OLE_CANCELBOOL readonly dispid 219;
    property esPostAceptacion: OLE_CANCELBOOL readonly dispid 220;
    property idComprobanteAsociado: IIdComprobanteTy readonly dispid 221;
    property referenciasComerciales: WideString readonly dispid 222;
    property arraySubtotalesIVA[Indice: Integer]: ISubtotalIVATy readonly dispid 223;
    property arraySubtotalesIVACount: Integer readonly dispid 224;
    property arrayOtrosTributos: IOtroTributoTy readonly dispid 225;
    property arrayOtrosTributosCount: Integer readonly dispid 226;
    property arrayItems: IItemTy readonly dispid 227;
    property arrayItemsCount: Integer readonly dispid 228;
    property datosGenerales: WideString readonly dispid 229;
    property datosComerciales: WideString readonly dispid 230;
    property leyendaComercial: WideString readonly dispid 231;
    property codCtaCte: Double readonly dispid 232;
    property estado_estado: WideString readonly dispid 233;
    property estado_fecha: WideString readonly dispid 234;
    property tipoAcep: WideString readonly dispid 235;
    property fechaHoraAcep: WideString readonly dispid 236;
    property arrayMotivosRechazo: IMotivoRechazoTy readonly dispid 237;
    property arrayMotivosRechazoCount: Integer readonly dispid 238;
    property infoAgDtpoCltv: OLE_CANCELBOOL readonly dispid 239;
    property fechaInfoAgDptoCltv: WideString readonly dispid 240;
    property idPagoAgDptoCltv: WideString readonly dispid 241;
    property CBUdePago: WideString readonly dispid 242;
  end;

// *********************************************************************//
// Interface: ISubtotalIVATy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {873012B5-0A40-440E-9F18-ED81C3C7AD4F}
// *********************************************************************//
  ISubtotalIVATy = interface(IDispatch)
    ['{873012B5-0A40-440E-9F18-ED81C3C7AD4F}']
  end;

// *********************************************************************//
// DispIntf:  ISubtotalIVATyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {873012B5-0A40-440E-9F18-ED81C3C7AD4F}
// *********************************************************************//
  ISubtotalIVATyDisp = dispinterface
    ['{873012B5-0A40-440E-9F18-ED81C3C7AD4F}']
  end;

// *********************************************************************//
// Interface: IOtroTributoTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {96443A17-3274-4493-A940-92F4FE8F4D98}
// *********************************************************************//
  IOtroTributoTy = interface(IDispatch)
    ['{96443A17-3274-4493-A940-92F4FE8F4D98}']
  end;

// *********************************************************************//
// DispIntf:  IOtroTributoTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {96443A17-3274-4493-A940-92F4FE8F4D98}
// *********************************************************************//
  IOtroTributoTyDisp = dispinterface
    ['{96443A17-3274-4493-A940-92F4FE8F4D98}']
  end;

// *********************************************************************//
// Interface: IItemTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {572B401B-91D9-46CA-85A7-ED286B14693B}
// *********************************************************************//
  IItemTy = interface(IDispatch)
    ['{572B401B-91D9-46CA-85A7-ED286B14693B}']
  end;

// *********************************************************************//
// DispIntf:  IItemTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {572B401B-91D9-46CA-85A7-ED286B14693B}
// *********************************************************************//
  IItemTyDisp = dispinterface
    ['{572B401B-91D9-46CA-85A7-ED286B14693B}']
  end;

// *********************************************************************//
// Interface: IMotivoRechazoTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {21AC85E6-B7A9-487F-BCBC-19E18AE05D42}
// *********************************************************************//
  IMotivoRechazoTy = interface(IDispatch)
    ['{21AC85E6-B7A9-487F-BCBC-19E18AE05D42}']
  end;

// *********************************************************************//
// DispIntf:  IMotivoRechazoTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {21AC85E6-B7A9-487F-BCBC-19E18AE05D42}
// *********************************************************************//
  IMotivoRechazoTyDisp = dispinterface
    ['{21AC85E6-B7A9-487F-BCBC-19E18AE05D42}']
  end;

// *********************************************************************//
// Interface: IInformarFacturaAgtDptoCltvRequestTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3DEF5AFE-1202-4B92-BEBB-4B006E29C02F}
// *********************************************************************//
  IInformarFacturaAgtDptoCltvRequestTy = interface(IDispatch)
    ['{3DEF5AFE-1202-4B92-BEBB-4B006E29C02F}']
    procedure idCtaCte(codCtaCte: Double); safecall;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); safecall;
    procedure ctaComitente(cuentaDepositante: Integer; subcuentaComitente: Double; 
                           const denominacion: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  IInformarFacturaAgtDptoCltvRequestTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3DEF5AFE-1202-4B92-BEBB-4B006E29C02F}
// *********************************************************************//
  IInformarFacturaAgtDptoCltvRequestTyDisp = dispinterface
    ['{3DEF5AFE-1202-4B92-BEBB-4B006E29C02F}']
    procedure idCtaCte(codCtaCte: Double); dispid 201;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); dispid 202;
    procedure ctaComitente(cuentaDepositante: Integer; subcuentaComitente: Double; 
                           const denominacion: WideString); dispid 203;
  end;

// *********************************************************************//
// Interface: IRechazarFECredRequestTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {30EBD9FB-D607-484D-A5E8-8AD7522DA407}
// *********************************************************************//
  IRechazarFECredRequestTy = interface(IDispatch)
    ['{30EBD9FB-D607-484D-A5E8-8AD7522DA407}']
    procedure idCtaCte(codCtaCte: Double); safecall;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); safecall;
    procedure arrayMotivosRechazo(codMotivo: Integer; const descMotivo: WideString; 
                                  const justificacion: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  IRechazarFECredRequestTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {30EBD9FB-D607-484D-A5E8-8AD7522DA407}
// *********************************************************************//
  IRechazarFECredRequestTyDisp = dispinterface
    ['{30EBD9FB-D607-484D-A5E8-8AD7522DA407}']
    procedure idCtaCte(codCtaCte: Double); dispid 201;
    procedure idFactura(CuitEmisor: Double; codTipoCmp: Integer; PtoVta: Integer; nroCmp: Double); dispid 202;
    procedure arrayMotivosRechazo(codMotivo: Integer; const descMotivo: WideString; 
                                  const justificacion: WideString); dispid 203;
  end;

// *********************************************************************//
// Interface: IconsultarObligadoRecepcionReturnTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2C7111F1-8465-43EB-9110-9303B3961AC3}
// *********************************************************************//
  IconsultarObligadoRecepcionReturnTy = interface(IDispatch)
    ['{2C7111F1-8465-43EB-9110-9303B3961AC3}']
    function Get_respuesta: OLE_CANCELBOOL; safecall;
    function Get_desde: WideString; safecall;
    property respuesta: OLE_CANCELBOOL read Get_respuesta;
    property desde: WideString read Get_desde;
  end;

// *********************************************************************//
// DispIntf:  IconsultarObligadoRecepcionReturnTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2C7111F1-8465-43EB-9110-9303B3961AC3}
// *********************************************************************//
  IconsultarObligadoRecepcionReturnTyDisp = dispinterface
    ['{2C7111F1-8465-43EB-9110-9303B3961AC3}']
    property respuesta: OLE_CANCELBOOL readonly dispid 201;
    property desde: WideString readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IConsultarMontoObligadoRecepcionReturnTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {24CDB620-0B79-4E7D-943A-3F55F1E26C95}
// *********************************************************************//
  IConsultarMontoObligadoRecepcionReturnTy = interface(IDispatch)
    ['{24CDB620-0B79-4E7D-943A-3F55F1E26C95}']
    function Get_obligado: OLE_CANCELBOOL; safecall;
    function Get_montoDesde: Double; safecall;
    property obligado: OLE_CANCELBOOL read Get_obligado;
    property montoDesde: Double read Get_montoDesde;
  end;

// *********************************************************************//
// DispIntf:  IConsultarMontoObligadoRecepcionReturnTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {24CDB620-0B79-4E7D-943A-3F55F1E26C95}
// *********************************************************************//
  IConsultarMontoObligadoRecepcionReturnTyDisp = dispinterface
    ['{24CDB620-0B79-4E7D-943A-3F55F1E26C95}']
    property obligado: OLE_CANCELBOOL readonly dispid 201;
    property montoDesde: Double readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IConsultarCtasCtesReturnTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9E84530B-FB93-4225-BB57-8BA22738ED6A}
// *********************************************************************//
  IConsultarCtasCtesReturnTy = interface(IDispatch)
    ['{9E84530B-FB93-4225-BB57-8BA22738ED6A}']
    function Get_arrayInfosCtaCte(Indice: Integer): IInfoCtaCteTy; safecall;
    function Get_arrayInfosCtaCteCount: Integer; safecall;
    property arrayInfosCtaCte[Indice: Integer]: IInfoCtaCteTy read Get_arrayInfosCtaCte;
    property arrayInfosCtaCteCount: Integer read Get_arrayInfosCtaCteCount;
  end;

// *********************************************************************//
// DispIntf:  IConsultarCtasCtesReturnTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9E84530B-FB93-4225-BB57-8BA22738ED6A}
// *********************************************************************//
  IConsultarCtasCtesReturnTyDisp = dispinterface
    ['{9E84530B-FB93-4225-BB57-8BA22738ED6A}']
    property arrayInfosCtaCte[Indice: Integer]: IInfoCtaCteTy readonly dispid 201;
    property arrayInfosCtaCteCount: Integer readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IInfoCtaCteTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AF2B653F-D0F7-40CA-BAF7-B8A30A2F03E0}
// *********************************************************************//
  IInfoCtaCteTy = interface(IDispatch)
    ['{AF2B653F-D0F7-40CA-BAF7-B8A30A2F03E0}']
    function Get_codCtaCte: Double; safecall;
    function Get_estadoCtaCte_estado: WideString; safecall;
    function Get_estadoCtaCte_fechaHoraEstado: WideString; safecall;
    function Get_CuitEmisor: Double; safecall;
    function Get_codTipoCmp: Integer; safecall;
    function Get_PtoVta: Integer; safecall;
    function Get_nroCmp: Double; safecall;
    function Get_importeTotalFC: Double; safecall;
    function Get_saldo: Double; safecall;
    function Get_saldoAceptado: Double; safecall;
    function Get_codMoneda: WideString; safecall;
    property codCtaCte: Double read Get_codCtaCte;
    property estadoCtaCte_estado: WideString read Get_estadoCtaCte_estado;
    property estadoCtaCte_fechaHoraEstado: WideString read Get_estadoCtaCte_fechaHoraEstado;
    property CuitEmisor: Double read Get_CuitEmisor;
    property codTipoCmp: Integer read Get_codTipoCmp;
    property PtoVta: Integer read Get_PtoVta;
    property nroCmp: Double read Get_nroCmp;
    property importeTotalFC: Double read Get_importeTotalFC;
    property saldo: Double read Get_saldo;
    property saldoAceptado: Double read Get_saldoAceptado;
    property codMoneda: WideString read Get_codMoneda;
  end;

// *********************************************************************//
// DispIntf:  IInfoCtaCteTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AF2B653F-D0F7-40CA-BAF7-B8A30A2F03E0}
// *********************************************************************//
  IInfoCtaCteTyDisp = dispinterface
    ['{AF2B653F-D0F7-40CA-BAF7-B8A30A2F03E0}']
    property codCtaCte: Double readonly dispid 201;
    property estadoCtaCte_estado: WideString readonly dispid 202;
    property estadoCtaCte_fechaHoraEstado: WideString readonly dispid 211;
    property CuitEmisor: Double readonly dispid 203;
    property codTipoCmp: Integer readonly dispid 204;
    property PtoVta: Integer readonly dispid 205;
    property nroCmp: Double readonly dispid 206;
    property importeTotalFC: Double readonly dispid 207;
    property saldo: Double readonly dispid 208;
    property saldoAceptado: Double readonly dispid 209;
    property codMoneda: WideString readonly dispid 210;
  end;

// *********************************************************************//
// Interface: IConsultarCtaCteReturnTy
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A70570EB-65D9-4117-A7AB-A57B902E3407}
// *********************************************************************//
  IConsultarCtaCteReturnTy = interface(IDispatch)
    ['{A70570EB-65D9-4117-A7AB-A57B902E3407}']
    function Get_codCtaCte: Integer; safecall;
    function Get_estadoCtaCte: WideString; safecall;
    function Get_factura: IComprobanteTy; safecall;
    function Get_importeInicial: Double; safecall;
    function Get_importeTotalNotasDC: Double; safecall;
    function Get_importeCancelado: Double; safecall;
    function Get_importeTotalRetPesos: Double; safecall;
    function Get_importeEmbargoPesos: Double; safecall;
    function Get_saldoAceptado: Double; safecall;
    function Get_saldo: Double; safecall;
    function Get_codMoneda: WideString; safecall;
    function Get_cotizacionMonedaUlt: Double; safecall;
    property codCtaCte: Integer read Get_codCtaCte;
    property estadoCtaCte: WideString read Get_estadoCtaCte;
    property factura: IComprobanteTy read Get_factura;
    property importeInicial: Double read Get_importeInicial;
    property importeTotalNotasDC: Double read Get_importeTotalNotasDC;
    property importeCancelado: Double read Get_importeCancelado;
    property importeTotalRetPesos: Double read Get_importeTotalRetPesos;
    property importeEmbargoPesos: Double read Get_importeEmbargoPesos;
    property saldoAceptado: Double read Get_saldoAceptado;
    property saldo: Double read Get_saldo;
    property codMoneda: WideString read Get_codMoneda;
    property cotizacionMonedaUlt: Double read Get_cotizacionMonedaUlt;
  end;

// *********************************************************************//
// DispIntf:  IConsultarCtaCteReturnTyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A70570EB-65D9-4117-A7AB-A57B902E3407}
// *********************************************************************//
  IConsultarCtaCteReturnTyDisp = dispinterface
    ['{A70570EB-65D9-4117-A7AB-A57B902E3407}']
    property codCtaCte: Integer readonly dispid 201;
    property estadoCtaCte: WideString readonly dispid 202;
    property factura: IComprobanteTy readonly dispid 203;
    property importeInicial: Double readonly dispid 204;
    property importeTotalNotasDC: Double readonly dispid 205;
    property importeCancelado: Double readonly dispid 206;
    property importeTotalRetPesos: Double readonly dispid 207;
    property importeEmbargoPesos: Double readonly dispid 208;
    property saldoAceptado: Double readonly dispid 209;
    property saldo: Double readonly dispid 210;
    property codMoneda: WideString readonly dispid 211;
    property cotizacionMonedaUlt: Double readonly dispid 212;
  end;

// *********************************************************************//
// Interface: IFEGenerador
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {429B1793-8E1A-4170-AFA9-7E499F3F6076}
// *********************************************************************//
  IFEGenerador = interface(IDispatch)
    ['{429B1793-8E1A-4170-AFA9-7E499F3F6076}']
    function Get_AbrirAlGenerar: OLE_CANCELBOOL; safecall;
    procedure Set_AbrirAlGenerar(Value: OLE_CANCELBOOL); safecall;
    function Editar(const Template: WideString): OLE_CANCELBOOL; safecall;
    function EnviarPorMail(const Template: WideString; const Email: WideString): OLE_CANCELBOOL; safecall;
    function Generar(const Archivo: WideString; const Template: WideString; 
                     const Formato: WideString): OLE_CANCELBOOL; safecall;
    procedure InsertaDetalle(const CodigoArt: WideString; const DescripcionArt: WideString; 
                             cantidad: Double; Precio: Double; Iva: Double; SubTotal: Double; 
                             const Info1: WideString; const Info2: WideString; 
                             const Info3: WideString; const Info4: WideString; 
                             const Info5: WideString; Num1: Double; Num2: Double; Num3: Double; 
                             Num4: Double; Num5: Double); safecall;
    procedure SeteaCabecera(TipoComp: Integer; PtoVta: Integer; Nro: Integer; 
                            const FechaComp: WideString; const RazonSocial: WideString; 
                            TipoDoc: Integer; const Documento: WideString; 
                            const CondIVA: WideString; const Cae: WideString; 
                            const Vencimiento: WideString; const direccion: WideString; 
                            Total: Double; const Info1: WideString; const Info2: WideString; 
                            const Info3: WideString; const Info4: WideString; 
                            const Info5: WideString); safecall;
    procedure SeteaCabeceraExtras(Num1: Double; Num2: Double; Num3: Double; Num4: Double; 
                                  Num5: Double; Num6: Double; Num7: Double; Num8: Double; 
                                  Num9: Double; Num10: Double; const Fecha1: WideString; 
                                  const Fecha2: WideString; const Fecha3: WideString; 
                                  const Fecha4: WideString; const Fecha5: WideString); safecall;
    procedure SeteaDatosVendedor(const CUIT: WideString; const RazonSocial: WideString; 
                                 const direccion: WideString; const Telefono: WideString; 
                                 const IngresosBrutos: WideString; const Info1: WideString; 
                                 const Info2: WideString; const Info3: WideString; 
                                 const Info4: WideString; const Info5: WideString); safecall;
    function Get_RutaLibreria: WideString; safecall;
    function Get_ErrorDesc: WideString; safecall;
    procedure SeteaCabeceraExtras2(const Moneda: WideString; Cotizacion: Double; 
                                   const TipoCodAut: WideString; const Info6: WideString; 
                                   const Info7: WideString; const Info8: WideString; 
                                   const Info9: WideString; const Info10: WideString); safecall;
    property AbrirAlGenerar: OLE_CANCELBOOL read Get_AbrirAlGenerar write Set_AbrirAlGenerar;
    property RutaLibreria: WideString read Get_RutaLibreria;
    property ErrorDesc: WideString read Get_ErrorDesc;
  end;

// *********************************************************************//
// DispIntf:  IFEGeneradorDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {429B1793-8E1A-4170-AFA9-7E499F3F6076}
// *********************************************************************//
  IFEGeneradorDisp = dispinterface
    ['{429B1793-8E1A-4170-AFA9-7E499F3F6076}']
    property AbrirAlGenerar: OLE_CANCELBOOL dispid 201;
    function Editar(const Template: WideString): OLE_CANCELBOOL; dispid 202;
    function EnviarPorMail(const Template: WideString; const Email: WideString): OLE_CANCELBOOL; dispid 203;
    function Generar(const Archivo: WideString; const Template: WideString; 
                     const Formato: WideString): OLE_CANCELBOOL; dispid 204;
    procedure InsertaDetalle(const CodigoArt: WideString; const DescripcionArt: WideString; 
                             cantidad: Double; Precio: Double; Iva: Double; SubTotal: Double; 
                             const Info1: WideString; const Info2: WideString; 
                             const Info3: WideString; const Info4: WideString; 
                             const Info5: WideString; Num1: Double; Num2: Double; Num3: Double; 
                             Num4: Double; Num5: Double); dispid 205;
    procedure SeteaCabecera(TipoComp: Integer; PtoVta: Integer; Nro: Integer; 
                            const FechaComp: WideString; const RazonSocial: WideString; 
                            TipoDoc: Integer; const Documento: WideString; 
                            const CondIVA: WideString; const Cae: WideString; 
                            const Vencimiento: WideString; const direccion: WideString; 
                            Total: Double; const Info1: WideString; const Info2: WideString; 
                            const Info3: WideString; const Info4: WideString; 
                            const Info5: WideString); dispid 206;
    procedure SeteaCabeceraExtras(Num1: Double; Num2: Double; Num3: Double; Num4: Double; 
                                  Num5: Double; Num6: Double; Num7: Double; Num8: Double; 
                                  Num9: Double; Num10: Double; const Fecha1: WideString; 
                                  const Fecha2: WideString; const Fecha3: WideString; 
                                  const Fecha4: WideString; const Fecha5: WideString); dispid 207;
    procedure SeteaDatosVendedor(const CUIT: WideString; const RazonSocial: WideString; 
                                 const direccion: WideString; const Telefono: WideString; 
                                 const IngresosBrutos: WideString; const Info1: WideString; 
                                 const Info2: WideString; const Info3: WideString; 
                                 const Info4: WideString; const Info5: WideString); dispid 208;
    property RutaLibreria: WideString readonly dispid 209;
    property ErrorDesc: WideString readonly dispid 210;
    procedure SeteaCabeceraExtras2(const Moneda: WideString; Cotizacion: Double; 
                                   const TipoCodAut: WideString; const Info6: WideString; 
                                   const Info7: WideString; const Info8: WideString; 
                                   const Info9: WideString; const Info10: WideString); dispid 211;
  end;

// *********************************************************************//
// Interface: IQr
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {75D6A95C-92FF-4A01-AD58-0AF819349713}
// *********************************************************************//
  IQr = interface(IDispatch)
    ['{75D6A95C-92FF-4A01-AD58-0AF819349713}']
    function Generar(ver: Integer; const fecha: WideString; CUIT: Double; PtoVta: Integer; 
                     tipoCmp: Integer; nroCmp: Double; Importe: Double; const Moneda: WideString; 
                     ctz: Double; tipoDocRec: Integer; nroDocRec: Double; 
                     const TipoCodAut: WideString; codAut: Double): OLE_CANCELBOOL; safecall;
    function Get_TextoQR: WideString; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_ArchivoQR: WideString; safecall;
    procedure Set_ArchivoQR(const Value: WideString); safecall;
    function Get_RutaLibreria: WideString; safecall;
    function Get_Base64: WideString; safecall;
    function Get_ArchivoPNG: WideString; safecall;
    procedure Set_ArchivoPNG(const Value: WideString); safecall;
    function GenerarTextoQR(ver: Integer; const fecha: WideString; CUIT: Double; PtoVta: Integer; 
                            tipoCmp: Integer; nroCmp: Double; Importe: Double; 
                            const Moneda: WideString; ctz: Double; tipoDocRec: Integer; 
                            nroDocRec: Double; const TipoCodAut: WideString; codAut: Double): OLE_CANCELBOOL; safecall;
    property TextoQR: WideString read Get_TextoQR;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property ArchivoQR: WideString read Get_ArchivoQR write Set_ArchivoQR;
    property RutaLibreria: WideString read Get_RutaLibreria;
    property Base64: WideString read Get_Base64;
    property ArchivoPNG: WideString read Get_ArchivoPNG write Set_ArchivoPNG;
  end;

// *********************************************************************//
// DispIntf:  IQrDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {75D6A95C-92FF-4A01-AD58-0AF819349713}
// *********************************************************************//
  IQrDisp = dispinterface
    ['{75D6A95C-92FF-4A01-AD58-0AF819349713}']
    function Generar(ver: Integer; const fecha: WideString; CUIT: Double; PtoVta: Integer; 
                     tipoCmp: Integer; nroCmp: Double; Importe: Double; const Moneda: WideString; 
                     ctz: Double; tipoDocRec: Integer; nroDocRec: Double; 
                     const TipoCodAut: WideString; codAut: Double): OLE_CANCELBOOL; dispid 225;
    property TextoQR: WideString readonly dispid 227;
    property ErrorDesc: WideString readonly dispid 201;
    property ArchivoQR: WideString dispid 202;
    property RutaLibreria: WideString readonly dispid 203;
    property Base64: WideString readonly dispid 204;
    property ArchivoPNG: WideString dispid 205;
    function GenerarTextoQR(ver: Integer; const fecha: WideString; CUIT: Double; PtoVta: Integer; 
                            tipoCmp: Integer; nroCmp: Double; Importe: Double; 
                            const Moneda: WideString; ctz: Double; tipoDocRec: Integer; 
                            nroDocRec: Double; const TipoCodAut: WideString; codAut: Double): OLE_CANCELBOOL; dispid 206;
  end;

// *********************************************************************//
// The Class Cowsaa provides a Create and CreateRemote method to          
// create instances of the default interface Iwsaa exposed by              
// the CoClass wsaa. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsaa = class
    class function Create: Iwsaa;
    class function CreateRemote(const MachineName: string): Iwsaa;
  end;

// *********************************************************************//
// The Class Cowsfexv1 provides a Create and CreateRemote method to          
// create instances of the default interface Iwsfexv1 exposed by              
// the CoClass wsfexv1. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsfexv1 = class
    class function Create: Iwsfexv1;
    class function CreateRemote(const MachineName: string): Iwsfexv1;
  end;

// *********************************************************************//
// The Class Cowsfev1 provides a Create and CreateRemote method to          
// create instances of the default interface Iwsfev1 exposed by              
// the CoClass wsfev1. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsfev1 = class
    class function Create: Iwsfev1;
    class function CreateRemote(const MachineName: string): Iwsfev1;
  end;

// *********************************************************************//
// The Class Cowsbfev1 provides a Create and CreateRemote method to          
// create instances of the default interface Iwsbfev1 exposed by              
// the CoClass wsbfev1. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsbfev1 = class
    class function Create: Iwsbfev1;
    class function CreateRemote(const MachineName: string): Iwsbfev1;
  end;

// *********************************************************************//
// The Class Cowsmtxca provides a Create and CreateRemote method to          
// create instances of the default interface Iwsmtxca exposed by              
// the CoClass wsmtxca. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsmtxca = class
    class function Create: Iwsmtxca;
    class function CreateRemote(const MachineName: string): Iwsmtxca;
  end;

// *********************************************************************//
// The Class Cowsseg provides a Create and CreateRemote method to          
// create instances of the default interface Iwsseg exposed by              
// the CoClass wsseg. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsseg = class
    class function Create: Iwsseg;
    class function CreateRemote(const MachineName: string): Iwsseg;
  end;

// *********************************************************************//
// The Class CowsPadron provides a Create and CreateRemote method to          
// create instances of the default interface IwsPadron exposed by              
// the CoClass wsPadron. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CowsPadron = class
    class function Create: IwsPadron;
    class function CreateRemote(const MachineName: string): IwsPadron;
  end;

// *********************************************************************//
// The Class CoComprobante provides a Create and CreateRemote method to          
// create instances of the default interface IComprobante exposed by              
// the CoClass Comprobante. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoComprobante = class
    class function Create: IComprobante;
    class function CreateRemote(const MachineName: string): IComprobante;
  end;

// *********************************************************************//
// The Class CoCbteAsoc provides a Create and CreateRemote method to          
// create instances of the default interface ICbteAsoc exposed by              
// the CoClass CbteAsoc. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCbteAsoc = class
    class function Create: ICbteAsoc;
    class function CreateRemote(const MachineName: string): ICbteAsoc;
  end;

// *********************************************************************//
// The Class CoTributo provides a Create and CreateRemote method to          
// create instances of the default interface ITributo exposed by              
// the CoClass Tributo. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTributo = class
    class function Create: ITributo;
    class function CreateRemote(const MachineName: string): ITributo;
  end;

// *********************************************************************//
// The Class CoAlicIva provides a Create and CreateRemote method to          
// create instances of the default interface IAlicIva exposed by              
// the CoClass AlicIva. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoAlicIva = class
    class function Create: IAlicIva;
    class function CreateRemote(const MachineName: string): IAlicIva;
  end;

// *********************************************************************//
// The Class CoOpcional provides a Create and CreateRemote method to          
// create instances of the default interface IOpcional exposed by              
// the CoClass Opcional. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOpcional = class
    class function Create: IOpcional;
    class function CreateRemote(const MachineName: string): IOpcional;
  end;

// *********************************************************************//
// The Class CoObs provides a Create and CreateRemote method to          
// create instances of the default interface IObs exposed by              
// the CoClass Obs. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoObs = class
    class function Create: IObs;
    class function CreateRemote(const MachineName: string): IObs;
  end;

// *********************************************************************//
// The Class CoContribuyente provides a Create and CreateRemote method to          
// create instances of the default interface IContribuyente exposed by              
// the CoClass Contribuyente. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoContribuyente = class
    class function Create: IContribuyente;
    class function CreateRemote(const MachineName: string): IContribuyente;
  end;

// *********************************************************************//
// The Class CoDomicilio provides a Create and CreateRemote method to          
// create instances of the default interface IDomicilio exposed by              
// the CoClass Domicilio. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDomicilio = class
    class function Create: IDomicilio;
    class function CreateRemote(const MachineName: string): IDomicilio;
  end;

// *********************************************************************//
// The Class CowsPadronARBA provides a Create and CreateRemote method to          
// create instances of the default interface IwsPadronARBA exposed by              
// the CoClass wsPadronARBA. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CowsPadronARBA = class
    class function Create: IwsPadronARBA;
    class function CreateRemote(const MachineName: string): IwsPadronARBA;
  end;

// *********************************************************************//
// The Class CoConsultaAlicuotaRespuesta provides a Create and CreateRemote method to          
// create instances of the default interface IConsultaAlicuotaRespuesta exposed by              
// the CoClass ConsultaAlicuotaRespuesta. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultaAlicuotaRespuesta = class
    class function Create: IConsultaAlicuotaRespuesta;
    class function CreateRemote(const MachineName: string): IConsultaAlicuotaRespuesta;
  end;

// *********************************************************************//
// The Class CoCertificado provides a Create and CreateRemote method to          
// create instances of the default interface ICertificado exposed by              
// the CoClass Certificado. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCertificado = class
    class function Create: ICertificado;
    class function CreateRemote(const MachineName: string): ICertificado;
  end;

// *********************************************************************//
// The Class Cowscdc provides a Create and CreateRemote method to          
// create instances of the default interface Iwscdc exposed by              
// the CoClass wscdc. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowscdc = class
    class function Create: Iwscdc;
    class function CreateRemote(const MachineName: string): Iwscdc;
  end;

// *********************************************************************//
// The Class CoBarcode provides a Create and CreateRemote method to          
// create instances of the default interface IBarcode exposed by              
// the CoClass Barcode. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoBarcode = class
    class function Create: IBarcode;
    class function CreateRemote(const MachineName: string): IBarcode;
  end;

// *********************************************************************//
// The Class Cowsct provides a Create and CreateRemote method to          
// create instances of the default interface Iwsct exposed by              
// the CoClass wsct. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsct = class
    class function Create: Iwsct;
    class function CreateRemote(const MachineName: string): Iwsct;
  end;

// *********************************************************************//
// The Class Cowsfecred provides a Create and CreateRemote method to          
// create instances of the default interface Iwsfecred exposed by              
// the CoClass wsfecred. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Cowsfecred = class
    class function Create: Iwsfecred;
    class function CreateRemote(const MachineName: string): Iwsfecred;
  end;

// *********************************************************************//
// The Class CoIdCtaCteTy provides a Create and CreateRemote method to          
// create instances of the default interface IIdCtaCteTy exposed by              
// the CoClass IdCtaCteTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoIdCtaCteTy = class
    class function Create: IIdCtaCteTy;
    class function CreateRemote(const MachineName: string): IIdCtaCteTy;
  end;

// *********************************************************************//
// The Class CoIdComprobanteTy provides a Create and CreateRemote method to          
// create instances of the default interface IIdComprobanteTy exposed by              
// the CoClass IdComprobanteTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoIdComprobanteTy = class
    class function Create: IIdComprobanteTy;
    class function CreateRemote(const MachineName: string): IIdComprobanteTy;
  end;

// *********************************************************************//
// The Class CoAceptarFECredRequestTy provides a Create and CreateRemote method to          
// create instances of the default interface IAceptarFECredRequestTy exposed by              
// the CoClass AceptarFECredRequestTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoAceptarFECredRequestTy = class
    class function Create: IAceptarFECredRequestTy;
    class function CreateRemote(const MachineName: string): IAceptarFECredRequestTy;
  end;

// *********************************************************************//
// The Class CoConsultarCmpReturnTy provides a Create and CreateRemote method to          
// create instances of the default interface IConsultarCmpReturnTy exposed by              
// the CoClass ConsultarCmpReturnTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultarCmpReturnTy = class
    class function Create: IConsultarCmpReturnTy;
    class function CreateRemote(const MachineName: string): IConsultarCmpReturnTy;
  end;

// *********************************************************************//
// The Class CoComprobanteTy provides a Create and CreateRemote method to          
// create instances of the default interface IComprobanteTy exposed by              
// the CoClass ComprobanteTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoComprobanteTy = class
    class function Create: IComprobanteTy;
    class function CreateRemote(const MachineName: string): IComprobanteTy;
  end;

// *********************************************************************//
// The Class CoSubtotalIVATy provides a Create and CreateRemote method to          
// create instances of the default interface ISubtotalIVATy exposed by              
// the CoClass SubtotalIVATy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoSubtotalIVATy = class
    class function Create: ISubtotalIVATy;
    class function CreateRemote(const MachineName: string): ISubtotalIVATy;
  end;

// *********************************************************************//
// The Class CoOtroTributoTy provides a Create and CreateRemote method to          
// create instances of the default interface IOtroTributoTy exposed by              
// the CoClass OtroTributoTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOtroTributoTy = class
    class function Create: IOtroTributoTy;
    class function CreateRemote(const MachineName: string): IOtroTributoTy;
  end;

// *********************************************************************//
// The Class CoItemTy provides a Create and CreateRemote method to          
// create instances of the default interface IItemTy exposed by              
// the CoClass ItemTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoItemTy = class
    class function Create: IItemTy;
    class function CreateRemote(const MachineName: string): IItemTy;
  end;

// *********************************************************************//
// The Class CoMotivoRechazoType provides a Create and CreateRemote method to          
// create instances of the default interface IMotivoRechazoTy exposed by              
// the CoClass MotivoRechazoType. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMotivoRechazoType = class
    class function Create: IMotivoRechazoTy;
    class function CreateRemote(const MachineName: string): IMotivoRechazoTy;
  end;

// *********************************************************************//
// The Class CoInformarFacturaAgtDptoCltvRequestTy provides a Create and CreateRemote method to          
// create instances of the default interface IInformarFacturaAgtDptoCltvRequestTy exposed by              
// the CoClass InformarFacturaAgtDptoCltvRequestTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoInformarFacturaAgtDptoCltvRequestTy = class
    class function Create: IInformarFacturaAgtDptoCltvRequestTy;
    class function CreateRemote(const MachineName: string): IInformarFacturaAgtDptoCltvRequestTy;
  end;

// *********************************************************************//
// The Class CoRechazarFECredRequestTy provides a Create and CreateRemote method to          
// create instances of the default interface IRechazarFECredRequestTy exposed by              
// the CoClass RechazarFECredRequestTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoRechazarFECredRequestTy = class
    class function Create: IRechazarFECredRequestTy;
    class function CreateRemote(const MachineName: string): IRechazarFECredRequestTy;
  end;

// *********************************************************************//
// The Class CoconsultarObligadoRecepcionReturnTy provides a Create and CreateRemote method to          
// create instances of the default interface IconsultarObligadoRecepcionReturnTy exposed by              
// the CoClass consultarObligadoRecepcionReturnTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoconsultarObligadoRecepcionReturnTy = class
    class function Create: IconsultarObligadoRecepcionReturnTy;
    class function CreateRemote(const MachineName: string): IconsultarObligadoRecepcionReturnTy;
  end;

// *********************************************************************//
// The Class CoConsultarMontoObligadoRecepcionReturnTy provides a Create and CreateRemote method to          
// create instances of the default interface IConsultarMontoObligadoRecepcionReturnTy exposed by              
// the CoClass ConsultarMontoObligadoRecepcionReturnTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultarMontoObligadoRecepcionReturnTy = class
    class function Create: IConsultarMontoObligadoRecepcionReturnTy;
    class function CreateRemote(const MachineName: string): IConsultarMontoObligadoRecepcionReturnTy;
  end;

// *********************************************************************//
// The Class CoConsultarCtasCtesReturnTy provides a Create and CreateRemote method to          
// create instances of the default interface IConsultarCtasCtesReturnTy exposed by              
// the CoClass ConsultarCtasCtesReturnTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultarCtasCtesReturnTy = class
    class function Create: IConsultarCtasCtesReturnTy;
    class function CreateRemote(const MachineName: string): IConsultarCtasCtesReturnTy;
  end;

// *********************************************************************//
// The Class CoInfoCtaCteTy provides a Create and CreateRemote method to          
// create instances of the default interface IInfoCtaCteTy exposed by              
// the CoClass InfoCtaCteTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoInfoCtaCteTy = class
    class function Create: IInfoCtaCteTy;
    class function CreateRemote(const MachineName: string): IInfoCtaCteTy;
  end;

// *********************************************************************//
// The Class CoConsultarCtaCteReturnTy provides a Create and CreateRemote method to          
// create instances of the default interface IConsultarCtaCteReturnTy exposed by              
// the CoClass ConsultarCtaCteReturnTy. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultarCtaCteReturnTy = class
    class function Create: IConsultarCtaCteReturnTy;
    class function CreateRemote(const MachineName: string): IConsultarCtaCteReturnTy;
  end;

// *********************************************************************//
// The Class CoFEGenerador provides a Create and CreateRemote method to          
// create instances of the default interface IFEGenerador exposed by              
// the CoClass FEGenerador. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFEGenerador = class
    class function Create: IFEGenerador;
    class function CreateRemote(const MachineName: string): IFEGenerador;
  end;

// *********************************************************************//
// The Class CoQr provides a Create and CreateRemote method to          
// create instances of the default interface IQr exposed by              
// the CoClass Qr. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoQr = class
    class function Create: IQr;
    class function CreateRemote(const MachineName: string): IQr;
  end;

implementation

uses ComObj;

class function Cowsaa.Create: Iwsaa;
begin
  Result := CreateComObject(CLASS_wsaa) as Iwsaa;
end;

class function Cowsaa.CreateRemote(const MachineName: string): Iwsaa;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsaa) as Iwsaa;
end;

class function Cowsfexv1.Create: Iwsfexv1;
begin
  Result := CreateComObject(CLASS_wsfexv1) as Iwsfexv1;
end;

class function Cowsfexv1.CreateRemote(const MachineName: string): Iwsfexv1;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsfexv1) as Iwsfexv1;
end;

class function Cowsfev1.Create: Iwsfev1;
begin
  Result := CreateComObject(CLASS_wsfev1) as Iwsfev1;
end;

class function Cowsfev1.CreateRemote(const MachineName: string): Iwsfev1;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsfev1) as Iwsfev1;
end;

class function Cowsbfev1.Create: Iwsbfev1;
begin
  Result := CreateComObject(CLASS_wsbfev1) as Iwsbfev1;
end;

class function Cowsbfev1.CreateRemote(const MachineName: string): Iwsbfev1;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsbfev1) as Iwsbfev1;
end;

class function Cowsmtxca.Create: Iwsmtxca;
begin
  Result := CreateComObject(CLASS_wsmtxca) as Iwsmtxca;
end;

class function Cowsmtxca.CreateRemote(const MachineName: string): Iwsmtxca;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsmtxca) as Iwsmtxca;
end;

class function Cowsseg.Create: Iwsseg;
begin
  Result := CreateComObject(CLASS_wsseg) as Iwsseg;
end;

class function Cowsseg.CreateRemote(const MachineName: string): Iwsseg;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsseg) as Iwsseg;
end;

class function CowsPadron.Create: IwsPadron;
begin
  Result := CreateComObject(CLASS_wsPadron) as IwsPadron;
end;

class function CowsPadron.CreateRemote(const MachineName: string): IwsPadron;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsPadron) as IwsPadron;
end;

class function CoComprobante.Create: IComprobante;
begin
  Result := CreateComObject(CLASS_Comprobante) as IComprobante;
end;

class function CoComprobante.CreateRemote(const MachineName: string): IComprobante;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Comprobante) as IComprobante;
end;

class function CoCbteAsoc.Create: ICbteAsoc;
begin
  Result := CreateComObject(CLASS_CbteAsoc) as ICbteAsoc;
end;

class function CoCbteAsoc.CreateRemote(const MachineName: string): ICbteAsoc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CbteAsoc) as ICbteAsoc;
end;

class function CoTributo.Create: ITributo;
begin
  Result := CreateComObject(CLASS_Tributo) as ITributo;
end;

class function CoTributo.CreateRemote(const MachineName: string): ITributo;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Tributo) as ITributo;
end;

class function CoAlicIva.Create: IAlicIva;
begin
  Result := CreateComObject(CLASS_AlicIva) as IAlicIva;
end;

class function CoAlicIva.CreateRemote(const MachineName: string): IAlicIva;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_AlicIva) as IAlicIva;
end;

class function CoOpcional.Create: IOpcional;
begin
  Result := CreateComObject(CLASS_Opcional) as IOpcional;
end;

class function CoOpcional.CreateRemote(const MachineName: string): IOpcional;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Opcional) as IOpcional;
end;

class function CoObs.Create: IObs;
begin
  Result := CreateComObject(CLASS_Obs) as IObs;
end;

class function CoObs.CreateRemote(const MachineName: string): IObs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Obs) as IObs;
end;

class function CoContribuyente.Create: IContribuyente;
begin
  Result := CreateComObject(CLASS_Contribuyente) as IContribuyente;
end;

class function CoContribuyente.CreateRemote(const MachineName: string): IContribuyente;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Contribuyente) as IContribuyente;
end;

class function CoDomicilio.Create: IDomicilio;
begin
  Result := CreateComObject(CLASS_Domicilio) as IDomicilio;
end;

class function CoDomicilio.CreateRemote(const MachineName: string): IDomicilio;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Domicilio) as IDomicilio;
end;

class function CowsPadronARBA.Create: IwsPadronARBA;
begin
  Result := CreateComObject(CLASS_wsPadronARBA) as IwsPadronARBA;
end;

class function CowsPadronARBA.CreateRemote(const MachineName: string): IwsPadronARBA;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsPadronARBA) as IwsPadronARBA;
end;

class function CoConsultaAlicuotaRespuesta.Create: IConsultaAlicuotaRespuesta;
begin
  Result := CreateComObject(CLASS_ConsultaAlicuotaRespuesta) as IConsultaAlicuotaRespuesta;
end;

class function CoConsultaAlicuotaRespuesta.CreateRemote(const MachineName: string): IConsultaAlicuotaRespuesta;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultaAlicuotaRespuesta) as IConsultaAlicuotaRespuesta;
end;

class function CoCertificado.Create: ICertificado;
begin
  Result := CreateComObject(CLASS_Certificado) as ICertificado;
end;

class function CoCertificado.CreateRemote(const MachineName: string): ICertificado;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Certificado) as ICertificado;
end;

class function Cowscdc.Create: Iwscdc;
begin
  Result := CreateComObject(CLASS_wscdc) as Iwscdc;
end;

class function Cowscdc.CreateRemote(const MachineName: string): Iwscdc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wscdc) as Iwscdc;
end;

class function CoBarcode.Create: IBarcode;
begin
  Result := CreateComObject(CLASS_Barcode) as IBarcode;
end;

class function CoBarcode.CreateRemote(const MachineName: string): IBarcode;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Barcode) as IBarcode;
end;

class function Cowsct.Create: Iwsct;
begin
  Result := CreateComObject(CLASS_wsct) as Iwsct;
end;

class function Cowsct.CreateRemote(const MachineName: string): Iwsct;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsct) as Iwsct;
end;

class function Cowsfecred.Create: Iwsfecred;
begin
  Result := CreateComObject(CLASS_wsfecred) as Iwsfecred;
end;

class function Cowsfecred.CreateRemote(const MachineName: string): Iwsfecred;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_wsfecred) as Iwsfecred;
end;

class function CoIdCtaCteTy.Create: IIdCtaCteTy;
begin
  Result := CreateComObject(CLASS_IdCtaCteTy) as IIdCtaCteTy;
end;

class function CoIdCtaCteTy.CreateRemote(const MachineName: string): IIdCtaCteTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_IdCtaCteTy) as IIdCtaCteTy;
end;

class function CoIdComprobanteTy.Create: IIdComprobanteTy;
begin
  Result := CreateComObject(CLASS_IdComprobanteTy) as IIdComprobanteTy;
end;

class function CoIdComprobanteTy.CreateRemote(const MachineName: string): IIdComprobanteTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_IdComprobanteTy) as IIdComprobanteTy;
end;

class function CoAceptarFECredRequestTy.Create: IAceptarFECredRequestTy;
begin
  Result := CreateComObject(CLASS_AceptarFECredRequestTy) as IAceptarFECredRequestTy;
end;

class function CoAceptarFECredRequestTy.CreateRemote(const MachineName: string): IAceptarFECredRequestTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_AceptarFECredRequestTy) as IAceptarFECredRequestTy;
end;

class function CoConsultarCmpReturnTy.Create: IConsultarCmpReturnTy;
begin
  Result := CreateComObject(CLASS_ConsultarCmpReturnTy) as IConsultarCmpReturnTy;
end;

class function CoConsultarCmpReturnTy.CreateRemote(const MachineName: string): IConsultarCmpReturnTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultarCmpReturnTy) as IConsultarCmpReturnTy;
end;

class function CoComprobanteTy.Create: IComprobanteTy;
begin
  Result := CreateComObject(CLASS_ComprobanteTy) as IComprobanteTy;
end;

class function CoComprobanteTy.CreateRemote(const MachineName: string): IComprobanteTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ComprobanteTy) as IComprobanteTy;
end;

class function CoSubtotalIVATy.Create: ISubtotalIVATy;
begin
  Result := CreateComObject(CLASS_SubtotalIVATy) as ISubtotalIVATy;
end;

class function CoSubtotalIVATy.CreateRemote(const MachineName: string): ISubtotalIVATy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SubtotalIVATy) as ISubtotalIVATy;
end;

class function CoOtroTributoTy.Create: IOtroTributoTy;
begin
  Result := CreateComObject(CLASS_OtroTributoTy) as IOtroTributoTy;
end;

class function CoOtroTributoTy.CreateRemote(const MachineName: string): IOtroTributoTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OtroTributoTy) as IOtroTributoTy;
end;

class function CoItemTy.Create: IItemTy;
begin
  Result := CreateComObject(CLASS_ItemTy) as IItemTy;
end;

class function CoItemTy.CreateRemote(const MachineName: string): IItemTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ItemTy) as IItemTy;
end;

class function CoMotivoRechazoType.Create: IMotivoRechazoTy;
begin
  Result := CreateComObject(CLASS_MotivoRechazoType) as IMotivoRechazoTy;
end;

class function CoMotivoRechazoType.CreateRemote(const MachineName: string): IMotivoRechazoTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MotivoRechazoType) as IMotivoRechazoTy;
end;

class function CoInformarFacturaAgtDptoCltvRequestTy.Create: IInformarFacturaAgtDptoCltvRequestTy;
begin
  Result := CreateComObject(CLASS_InformarFacturaAgtDptoCltvRequestTy) as IInformarFacturaAgtDptoCltvRequestTy;
end;

class function CoInformarFacturaAgtDptoCltvRequestTy.CreateRemote(const MachineName: string): IInformarFacturaAgtDptoCltvRequestTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InformarFacturaAgtDptoCltvRequestTy) as IInformarFacturaAgtDptoCltvRequestTy;
end;

class function CoRechazarFECredRequestTy.Create: IRechazarFECredRequestTy;
begin
  Result := CreateComObject(CLASS_RechazarFECredRequestTy) as IRechazarFECredRequestTy;
end;

class function CoRechazarFECredRequestTy.CreateRemote(const MachineName: string): IRechazarFECredRequestTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_RechazarFECredRequestTy) as IRechazarFECredRequestTy;
end;

class function CoconsultarObligadoRecepcionReturnTy.Create: IconsultarObligadoRecepcionReturnTy;
begin
  Result := CreateComObject(CLASS_consultarObligadoRecepcionReturnTy) as IconsultarObligadoRecepcionReturnTy;
end;

class function CoconsultarObligadoRecepcionReturnTy.CreateRemote(const MachineName: string): IconsultarObligadoRecepcionReturnTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_consultarObligadoRecepcionReturnTy) as IconsultarObligadoRecepcionReturnTy;
end;

class function CoConsultarMontoObligadoRecepcionReturnTy.Create: IConsultarMontoObligadoRecepcionReturnTy;
begin
  Result := CreateComObject(CLASS_ConsultarMontoObligadoRecepcionReturnTy) as IConsultarMontoObligadoRecepcionReturnTy;
end;

class function CoConsultarMontoObligadoRecepcionReturnTy.CreateRemote(const MachineName: string): IConsultarMontoObligadoRecepcionReturnTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultarMontoObligadoRecepcionReturnTy) as IConsultarMontoObligadoRecepcionReturnTy;
end;

class function CoConsultarCtasCtesReturnTy.Create: IConsultarCtasCtesReturnTy;
begin
  Result := CreateComObject(CLASS_ConsultarCtasCtesReturnTy) as IConsultarCtasCtesReturnTy;
end;

class function CoConsultarCtasCtesReturnTy.CreateRemote(const MachineName: string): IConsultarCtasCtesReturnTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultarCtasCtesReturnTy) as IConsultarCtasCtesReturnTy;
end;

class function CoInfoCtaCteTy.Create: IInfoCtaCteTy;
begin
  Result := CreateComObject(CLASS_InfoCtaCteTy) as IInfoCtaCteTy;
end;

class function CoInfoCtaCteTy.CreateRemote(const MachineName: string): IInfoCtaCteTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InfoCtaCteTy) as IInfoCtaCteTy;
end;

class function CoConsultarCtaCteReturnTy.Create: IConsultarCtaCteReturnTy;
begin
  Result := CreateComObject(CLASS_ConsultarCtaCteReturnTy) as IConsultarCtaCteReturnTy;
end;

class function CoConsultarCtaCteReturnTy.CreateRemote(const MachineName: string): IConsultarCtaCteReturnTy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultarCtaCteReturnTy) as IConsultarCtaCteReturnTy;
end;

class function CoFEGenerador.Create: IFEGenerador;
begin
  Result := CreateComObject(CLASS_FEGenerador) as IFEGenerador;
end;

class function CoFEGenerador.CreateRemote(const MachineName: string): IFEGenerador;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FEGenerador) as IFEGenerador;
end;

class function CoQr.Create: IQr;
begin
  Result := CreateComObject(CLASS_Qr) as IQr;
end;

class function CoQr.CreateRemote(const MachineName: string): IQr;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Qr) as IQr;
end;

end.
