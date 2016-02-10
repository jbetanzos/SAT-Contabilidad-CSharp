using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SatContabilidad
{
    class Program
    {
        static void Main(string[] args)
        {
            var AppSettings = System.Configuration.ConfigurationManager.AppSettings;
            Vendor.Contabilidad contabilidad = new Vendor.Contabilidad(
                AppSettings["OutputFolder"], 
                AppSettings["VendorRFC"], 
                AppSettings["DatabaseFileName"],
                AppSettings["Month"],
                AppSettings["Year"]
                );

            contabilidad.createPolizasPeriodo();
            contabilidad.createCatalagoCuentas();
            contabilidad.createBalanza();

            Console.WriteLine("Press enter to continue...");
            Console.ReadLine();
        }
    }
}
