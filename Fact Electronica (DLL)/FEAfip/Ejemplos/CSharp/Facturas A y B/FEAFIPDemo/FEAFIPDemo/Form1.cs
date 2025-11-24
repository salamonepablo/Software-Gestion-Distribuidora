using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FEAFIPLib;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /* Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
            y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
            la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.*/
            const
                //URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
                string URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms";
            //Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
            string URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx";
            // Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx

            // Agregar FEAFIPlib como referencia al proyecto desde el menu y luego en el using.

            string CAE = "";
            string Vencimiento = "";
            string Resultado = "";
            string Reproceso = "";
            double nro = 0;
            int PtoVta = 10;
            int TipoComp = 1; // Factura A(ir a http://www.bitingenieria.com.ar/codigos.html)
            string FechaComp = DateTime.Today.ToString("yyyyMMdd");

            wsfev1 lwsfev1 = new wsfev1();
            lwsfev1.CUIT = 20939802593; // Cuit del vendedor
            lwsfev1.URL = URLWSW;
            if (lwsfev1.login("certificado.crt", "clave.key", URLWSAA))
            {
                if (lwsfev1.SFRecuperaLastCMP(PtoVta, TipoComp) == false)
                {
                    MessageBox.Show(lwsfev1.ErrorDesc);
                }
                else
                {
                    nro = lwsfev1.SFLastCMP + 1;
                    lwsfev1.Reset();
                    lwsfev1.AgregaFactura(1, 80, 30707219072, nro, nro, FechaComp, 121, 0, 100, 0, "", "", "", "PES", 1);
                    lwsfev1.AgregaIVA(5, 100, 21); //ir a http://www.bitingenieria.com.ar/codigos.html
                    if (lwsfev1.Autorizar(PtoVta, (int) FEAFIPLib.TipoComprobante.tcFacturaA))
                    {
                        lwsfev1.AutorizarRespuesta(0, out CAE, out Vencimiento, out Resultado, out Reproceso);
                        if (Resultado == "A")
                        {
                            MessageBox.Show("Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento "
                                + ":" + CAE + " " + Vencimiento);
                        }
                        else
                        {
                            MessageBox.Show(lwsfev1.AutorizarRespuestaObs(0));
                        }
                    }
                    else
                    {
                        MessageBox.Show(lwsfev1.ErrorDesc);
                    }
                }
            }
            else
            {
                MessageBox.Show(lwsfev1.ErrorDesc);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            wsPadron lwsPadron = new wsPadron(); 
            Contribuyente contribuyente = new Contribuyente();

            // Los datos de pruebas pueden ser obtenidos de aqui http://www.afip.gov.ar/ws/ws_sr_padron_a4/datos-prueba-padron-a4.txt
            lwsPadron.CUIT = 20939802593;
            lwsPadron.ModoProduccion = false; // Para poder consultar un CUIT real debe habilitar el modo producción.
            if (lwsPadron.login("certificado.crt", "clave.key")) {
                if (lwsPadron.consultar(30202020204, contribuyente))
                {
                    
                    string Datos = contribuyente.nombre;
                    Datos = Datos + "\r" +contribuyente.tipoPersona;
                    Datos = Datos + "\r" + contribuyente.domicilioFiscal.direccion;
                    Datos = Datos + "\r" + contribuyente.domicilioFiscal.provincia;
                    // Control de constancia de inscripción. Si hay observaciones el contribuyente tiene errores de constancia.
                    Datos = Datos + "\r" + contribuyente.observaciones;
                    MessageBox.Show(Datos);
                } else {
                    MessageBox.Show(lwsPadron.ErrorDesc);
               }
        } else {
                MessageBox.Show(lwsPadron.ErrorDesc);
    }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            wsPadronARBA lwsPadron = new wsPadronARBA(); 
            lwsPadron.User = "20939802593";
            lwsPadron.Password = "123456";
            lwsPadron.ModoProduccion = false; // Debe dar de alta el cuit en el entorno de test de ARBA http://www.test.arba.gov.ar/
            if (lwsPadron.ConsultaAlicuota("20160701", "20160731", 27929007862)) {
                string Datos = "Alícuota Percepción: " + lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaPercepcion.ToString();
                Datos = Datos + "\r" + "Alícuota Retención: " + lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaRetencion.ToString();
                MessageBox.Show(Datos);
            } else {
                MessageBox.Show(lwsPadron.ErrorDesc);
            }
        }

        
    }
}
