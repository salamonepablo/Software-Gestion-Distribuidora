using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FEAFIPLib;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        const string URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms";
        const string URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          int PtoVta = 2000; // Un Punto de Venta dado de alta en la AFIP;
          int TipoComp = 206; // Factura A(Ver excel referencias codigos AFIP documentacion.rar)
          double nro;
          string fecha = DateTime.Today.ToString("yyyyMMdd");
          string fechaVenc = DateTime.Today.AddDays(30).ToString("yyyyMMdd");
  
          wsfev1 lwsfev1 = new wsfev1();
          lwsfev1.CUIT = 20939802593;
          lwsfev1.URL = URLWSW;

          lwsfev1.Depurar = true;
          if (lwsfev1.login("certificado.crt", "clave.key", URLWSAA)) 
          {
              if (!lwsfev1.RecuperaLastCMP(PtoVta, TipoComp, out nro)) 
              {
                  MessageBox.Show(lwsfev1.ErrorDesc);
                  return;
              }
              nro = nro + 1;

              lwsfev1.AgregaFactura(1, 80, 30504507323, nro, nro,
                fecha, 7260000, 0, 6000000, 0,
                "", "", fechaVenc, "PES", 1);
              lwsfev1.AgregaIVA(5, 6000000, 1260000); // IVA 21
              lwsfev1.AgregaOpcional("2101", "0150507801000124703453");
              if (lwsfev1.Autorizar(PtoVta, TipoComp))
              {
                if (lwsfev1.SFResultado[0] == "A") 
                  MessageBox.Show("Felicitaciones! CAE y Vencimiento: " + lwsfev1.SFCAE[0] + "/" + lwsfev1.SFVencimiento[0]);
                else
                  MessageBox.Show(lwsfev1.AutorizarRespuestaObs(0));
              }
              else
                MessageBox.Show(lwsfev1.ErrorDesc);
          }
          else
              MessageBox.Show(lwsfev1.ErrorDesc);
        
        }

        private void button3_Click(object sender, EventArgs e)
        {
          double codCtaCte = 1;
          double CUITEmisor = 27929007862;
          int codTipoCmp = 201;
          int ptoVta = 10;
          double nroCmp = 1;

          wsfecred fceObj = new wsfecred();
          fceObj.CUIT = 20939802593;
          fceObj.ModoProduccion = false;
          fceObj.Depurar = true;
          if (fceObj.login("certificado.crt", "clave.key"))
          {
              InformarFacturaAgtDptoCltvRequestTy informarFECred = fceObj.nuevoInformarFacturaAgtDptoCltvRequestTy();
              //Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
              //informarFECred.idCtaCte(codCtaCte);
              informarFECred.idFactura(
                CUITEmisor,
                codTipoCmp,
                ptoVta,
                nroCmp);
              int cuentaDepositante = 250;
              double subcuentaComitente = 120310;
              string denominacion = "denominacion";
              informarFECred.ctaComitente(cuentaDepositante, subcuentaComitente, denominacion);
              if (fceObj.informarFacturaAgtDptoCltv(informarFECred))
                  MessageBox.Show("Operación realizada con éxito");
              else
                  MessageBox.Show(fceObj.ErrorDesc);
          }
          else
              MessageBox.Show(fceObj.ErrorDesc);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            long CUITEmisor = 27929007862;
            int codTipoCmp = 201;
            int ptoVta = 10;
            long nroCmp = 1;

            wsfecred fceObj = new wsfecred();
            fceObj.CUIT = 20939802593;
            fceObj.ModoProduccion = false;
            fceObj.Depurar = true;
            if (fceObj.login("certificado.crt", "clave.key"))
            {
                AceptarFECredRequestTy aceptarFECred = fceObj.nuevoAceptarFECredRequestTy();

                //Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
                //aceptarFECred.idCtaCte(codCtaCte);
                aceptarFECred.idFactura(
                  CUITEmisor,
                  codTipoCmp,
                  ptoVta,
                  nroCmp);
                aceptarFECred.tipoCancelacion = "TOT";
                aceptarFECred.importeCancelado = 10.1;
                aceptarFECred.importeTotalRetPesos = 11.2;
                aceptarFECred.importeEmbargoPesos = 12.4;
                aceptarFECred.saldoAceptado = 13.6;
                aceptarFECred.codMoneda = "PES";
                aceptarFECred.cotizacionMonedaUlt = 1.2;

                //Confirmar Notas 1..N
                bool cacepta = true;
                double cCUITEmisor = 27929007862;
                int ccodTipoCmp = 201;
                int cptoVta = 10;
                long cnroCmp = 1;

                aceptarFECred.arrayConfirmarNotasDC(
                  cacepta, cCUITEmisor, ccodTipoCmp, cptoVta, cnroCmp);

                //Formas de cancelacion 1..N
                int codigoCancelacion = 1;
                string descripcionCancelacion = "descripcionCancelacion";
                aceptarFECred.arrayFormasCancelacion(codigoCancelacion, descripcionCancelacion);

                //Retenciones 1..N
                int rcodTipo = 1;
                double rimporte = 10.3;
                double rporcentaje = 5.5;
                string rdescMotivo = "Motivo 2";
                aceptarFECred.arrayRetenciones(rcodTipo, rimporte, rporcentaje, rdescMotivo);
                rcodTipo = 1;
                rimporte = 0.9;
                rporcentaje = 0.0;
                rdescMotivo = "Motivo 2";
                aceptarFECred.arrayRetenciones(rcodTipo, rimporte, rporcentaje, rdescMotivo);

                //Ajustes 1..N
                //int acodigo = 1;
                //double aimporte = 10.3;
                //aceptarFECred.arrayAjustesOperacion(acodigo, aimporte);

                if (fceObj.aceptarFECred(aceptarFECred))
                    MessageBox.Show("Operación realizada con éxito");
                else
                    MessageBox.Show(fceObj.ErrorDesc);
            }
            else
                MessageBox.Show(fceObj.ErrorDesc);

        }

        private void button5_Click(object sender, EventArgs e)
        {
          long CUITEmisor = 27929007862;
          int codTipoCmp = 201;
          int ptoVta = 10;
          long nroCmp = 1;

          Iwsfecred fceObj = new wsfecred();
          fceObj.CUIT = 20939802593;
          fceObj.ModoProduccion = false;
          fceObj.Depurar = true;
          if (fceObj.login("certificado.crt", "clave.key")) {
          
             RechazarFECredRequestTy rechazarFECred = fceObj.nuevoRechazarFECredRequestTy();

            //Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
            //aceptarFECred.idCtaCte(codCtaCte);
            rechazarFECred.idFactura(
              CUITEmisor,
              codTipoCmp,
              ptoVta,
              nroCmp);

            int codMotivo = 1;
            string descMotivo = "Mercaderia dañada";
            string justificacion = "Accidente vial";

            rechazarFECred.arrayMotivosRechazo(codMotivo, descMotivo, justificacion);
            if (fceObj.rechazarFECred(rechazarFECred)) 
              MessageBox.Show("Operación realizada con éxito");
            else
              MessageBox.Show(fceObj.ErrorDesc);
          } 
          else
            MessageBox.Show(fceObj.ErrorDesc);
        }

        private void button6_Click(object sender, EventArgs e)
        {
          string rolCUITRepresentada = "RECEPTOR";
          double CUITContraparte = 0;
          int codTipoCmp = 0;
          string estadoCmp = "";
          string fecha_tipo = "";
          string fechaDesde = "";
          string fecha_hasta = "";
          int codCtaCte = 0;
          string estadoCtaCte = "";

          wsfecred fceObj = new wsfecred();
          fceObj.CUIT = 30504507323;
          fceObj.ModoProduccion = false;
          fceObj.Depurar = true;
          if (fceObj.login("certificado.crt", "clave.key")) 
          {
            if (fceObj.consultarComprobantes(rolCUITRepresentada,
              CUITContraparte, codTipoCmp, estadoCmp, fecha_tipo, fechaDesde, fecha_hasta,
              codCtaCte, estadoCtaCte))
            {
               ConsultarCmpReturnTy consultarComprobantesReturn = fceObj.consultarCmpReturn;
              MessageBox.Show("Consulta realizada con éxito. Cantidad de comprobantes: " + consultarComprobantesReturn.arrayComprobantesCount.ToString());
            }
            else
              MessageBox.Show(fceObj.ErrorDesc);
          }
          else
            MessageBox.Show(fceObj.ErrorDesc);
        }

        private void button2_Click(object sender, EventArgs e)
        {
          string rolCUITRepresentada = "EMISOR";
          double CUITContraparte = 0;
          int codTipoCmp = 0;
          string estadoCmp = "";
          string fecha_tipo = "";
          string fechaDesde = "";
          string fecha_hasta = "";
          int codCtaCte = 0;
          string estadoCtaCte = "";

          wsfecred fceObj = new wsfecred();
          fceObj.CUIT = 20939802593;
          fceObj.ModoProduccion = false;
          fceObj.Depurar = true;
          if (fceObj.login("certificado.crt", "clave.key")) 
          {
            if (fceObj.consultarComprobantes(rolCUITRepresentada,
              CUITContraparte, codTipoCmp, estadoCmp, fecha_tipo, fechaDesde, fecha_hasta,
              codCtaCte, estadoCtaCte)) 
            {
              ConsultarCmpReturnTy consultarComprobantesReturn = fceObj.consultarCmpReturn;
              MessageBox.Show("Consulta realizada con éxito. Cantidad de comprobantes: " + consultarComprobantesReturn.arrayComprobantesCount.ToString());
            }
            else
              MessageBox.Show(fceObj.ErrorDesc);
          }
          else
            MessageBox.Show(fceObj.ErrorDesc);

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
