using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportProduct
{
    class Program
    {
        static void Main(string[] args)
        {

        }
    }

    public class Product
    {
        public Guid Id { get; set; }

        public string CodeSap { get; set; }

        public string TurkName { get; set; }

        public string EngName { get; set; }

        public string RuName { get; set; }

        public string Type { get; set; }

        public string TypeEquipment { get; set; }

        public string Serial { get; set; }

        public int TcmTRKarExworksEuro { get; set; }

        public int TcmFiyatListesiEuro { get; set; }

        public int TcmFromStockToIstanbulEuro { get; set; }

        public int TcmTransportEuro { get; set; }

        public int TcmGeneralExpencesEuro { get; set; }

        public int TcmProfitEuro { get; set; }

        public int TcmIskontoPayEuro { get; set; }

        public int TcmNDSEuro { get; set; }
    }
}
