using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Xml.Serialization;

namespace SatContabilidad.Vendor
{
    class Contabilidad
    {
        private string outputPath;
        private string rfc;
        private string databaseFileName;
        private string month;
        private string year;

        public Contabilidad(string outputPath, string rfc, string databaseFileName, string month, string year)
        {
            this.outputPath = outputPath;
            this.rfc = rfc;
            this.databaseFileName = databaseFileName;
            this.month = month;
            this.year = year;
        }

        public void createBalanza()
        {
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database\\" + this.databaseFileName;
            string strAccessSelect = "SELECT * FROM [Add Acumulados] INNER JOIN [CatalogoCuentasSatVendor] ON [Add Acumulados].[SUBCTA] = [CatalogoCuentasSatVendor].[NumCta]";

            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataReader reader = myAccessCommand.ExecuteReader();
                List<SatContabilidad.Balanza.BalanzaCtas> balanzaCtas = new List<SatContabilidad.Balanza.BalanzaCtas>();

                while (reader.Read())
                {
                    SatContabilidad.Balanza.BalanzaCtas balanzaCta = new Balanza.BalanzaCtas();
                    balanzaCta.NumCta = reader["SUBCTA"].ToString().Trim();
                    balanzaCta.SaldoIni = 0;
                    balanzaCta.Debe = Decimal.Parse(reader["MOVDEB"].ToString());
                    balanzaCta.Haber = Decimal.Parse(reader["MOVHAB"].ToString());
                    balanzaCta.SaldoFin = balanzaCta.Haber - balanzaCta.Debe;
                    balanzaCtas.Add(balanzaCta);
                }

                reader.Close();
                myAccessConn.Close();

                SatContabilidad.Balanza.Balanza balanza = new Balanza.Balanza();
                balanza.TipoEnvio = "N";
                balanza.schemaLocation = "www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion http://www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion/BalanzaComprobacion_1_1.xsd";
                balanza.Anio = int.Parse(this.year);
                balanza.Ctas = balanzaCtas.ToArray<Balanza.BalanzaCtas>();
                Type mes = typeof(Balanza.BalanzaMes);
                
                balanza.Mes = (Balanza.BalanzaMes) Enum.Parse(mes, (int.Parse(this.month) - 1).ToString());
                balanza.RFC = this.rfc;

                XmlSerializer serializer = new XmlSerializer(typeof(SatContabilidad.Balanza.Balanza));
                System.IO.TextWriter writer = new System.IO.StreamWriter(this.outputPath + this.rfc + this.year + this.month + "BN.xml");

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("BCE", "www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion");
                ns.Add("xsi", "http://www.w3.org/2001/XMLSchema-instance");

                serializer.Serialize(writer, balanza, ns);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrive the required data from the Database.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }
        }

        public void createPolizasPeriodo()
        {
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database\\" + this.databaseFileName;
            string strAccessSelect = "SELECT * FROM [Apuntes_Polizas_Import]";

            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataReader reader = myAccessCommand.ExecuteReader();
                List<SatContabilidad.Polizas.PolizasPoliza> polizas = new List<SatContabilidad.Polizas.PolizasPoliza>();

                while (reader.Read())
                {
                    SatContabilidad.Polizas.PolizasPoliza poliza = new SatContabilidad.Polizas.PolizasPoliza();
                    List<SatContabilidad.Polizas.PolizasPolizaTransaccion> transacciones = new List<Polizas.PolizasPolizaTransaccion>();
                    SatContabilidad.Polizas.PolizasPolizaTransaccion transaccion = new Polizas.PolizasPolizaTransaccion();
                    List<SatContabilidad.Polizas.PolizasPolizaTransaccionOtrMetodoPago> otrosMetodosdePago = new List<Polizas.PolizasPolizaTransaccionOtrMetodoPago>();

                    Polizas.PolizasPolizaTransaccionOtrMetodoPago otroMetodoPago = new Polizas.PolizasPolizaTransaccionOtrMetodoPago();

                    switch (reader["MONEDA"].ToString())
                    {
                        case "DL":
                            otroMetodoPago.Moneda = SatContabilidad.Polizas.c_Moneda.USD;
                            otroMetodoPago.TipCamb = Decimal.Parse(reader["CAMAPU"].ToString());
                            otroMetodoPago.TipCambSpecified = true;
                            break;
                        default:
                            otroMetodoPago.Moneda = SatContabilidad.Polizas.c_Moneda.MXN;
                            otroMetodoPago.TipCambSpecified = false;
                            break;
                    }
                    
                    otroMetodoPago.MonedaSpecified = true;
                    otroMetodoPago.Monto = Decimal.Parse(reader["IMPAPU"].ToString());
                    otroMetodoPago.RFC = this.rfc;
                    

                    otrosMetodosdePago.Add(otroMetodoPago);
                    transaccion.OtrMetodoPago = otrosMetodosdePago.ToArray<SatContabilidad.Polizas.PolizasPolizaTransaccionOtrMetodoPago>();
                    transacciones.Add(transaccion);
                    poliza.Transaccion = transacciones.ToArray<SatContabilidad.Polizas.PolizasPolizaTransaccion>();
                    polizas.Add(poliza);
                }

                reader.Close();
                myAccessConn.Close();

                SatContabilidad.Polizas.Polizas polizasLast = new SatContabilidad.Polizas.Polizas();
                polizasLast.schemaLocation = "www.sat.gob.mx/esquemas/ContabilidadE/1_1/CatalogoCuentas/CatalogoCuentas_1_1.xsd";
                polizasLast.Anio = int.Parse(this.year);
                Type mes = typeof(Polizas.PolizasMes);
                polizasLast.Mes = (Polizas.PolizasMes) Enum.Parse(mes, (int.Parse(this.month) - 1).ToString());
                polizasLast.Poliza = polizas.ToArray< SatContabilidad.Polizas.PolizasPoliza>();
                polizasLast.RFC = this.rfc;

                XmlSerializer serializer = new XmlSerializer(typeof(SatContabilidad.Polizas.Polizas));
                System.IO.TextWriter writer = new System.IO.StreamWriter(this.outputPath + this.rfc + this.year + this.month + "PL.xml");

                serializer.Serialize(writer, polizasLast);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrive the required data from the Database.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }
        }

        public void createCatalagoCuentas()
        {
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database\\" + this.databaseFileName;
            string strAccessSelect = "SELECT * FROM [CatalogoCuentasSatVendor]";

            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataReader reader = myAccessCommand.ExecuteReader();
                List<SatContabilidad.CatalogoCtas.CatalogoCtas> ctas = new List<SatContabilidad.CatalogoCtas.CatalogoCtas>();

                while (reader.Read())
                {
                    SatContabilidad.CatalogoCtas.CatalogoCtas cta = new SatContabilidad.CatalogoCtas.CatalogoCtas();

                    Type code = typeof(CatalogoCtas.c_CodAgrup);

                    foreach (CatalogoCtas.c_CodAgrup o in Enum.GetValues(typeof(CatalogoCtas.c_CodAgrup)))
                    {
                        if (GetXmlAttrNameFromEnumValue(o).Equals(reader["CodAgrupador"].ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            cta.CodAgrup = o;
                        }
                    }

                    cta.NumCta = reader["NumCta"].ToString();
                    cta.Desc = reader["Desc"].ToString();
                    cta.Natur = reader["Natur"].ToString();
                    cta.Nivel = int.Parse(reader["Nivel"].ToString());

                    ctas.Add(cta);
                }

                reader.Close();
                myAccessConn.Close();

                SatContabilidad.CatalogoCtas.Catalogo catalogo = new SatContabilidad.CatalogoCtas.Catalogo();
                catalogo.schemaLocation = "www.sat.gob.mx/esquemas/ContabilidadE/1_1/CatalogoCuentas http://www.sat.gob.mx/esquemas/ContabilidadE/1_1/CatalogoCuentas/CatalogoCuentas_1_1.xsd";
                catalogo.Ctas = ctas.ToArray<SatContabilidad.CatalogoCtas.CatalogoCtas>();
                Type mes = typeof(CatalogoCtas.CatalogoMes);
                
                catalogo.Mes = (CatalogoCtas.CatalogoMes)Enum.Parse(mes, (int.Parse(this.month) - 1).ToString());
                catalogo.Anio = int.Parse(this.year);
                catalogo.RFC = this.rfc;

                XmlSerializer serializer = new XmlSerializer(typeof(SatContabilidad.CatalogoCtas.Catalogo));
                System.IO.TextWriter writer = new System.IO.StreamWriter(this.outputPath + this.rfc + this.year + this.month + "CT.xml");

                serializer.Serialize(writer, catalogo);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrive the required data from the Database.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }
        }

        public string GetXmlAttrNameFromEnumValue(CatalogoCtas.c_CodAgrup pEnumVal)
        {
            Type type = pEnumVal.GetType();
            System.Reflection.FieldInfo info = type.GetField(Enum.GetName(typeof(CatalogoCtas.c_CodAgrup), pEnumVal));
            XmlEnumAttribute att = (XmlEnumAttribute)info.GetCustomAttributes(typeof(XmlEnumAttribute), false)[0];

            return att.Name;
        }
    }
}
