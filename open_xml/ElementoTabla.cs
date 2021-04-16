using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace open_xml
{
    class ElementoTabla
    {
        public string Codigo { get; set; }
                
        public int Cantidad { get; set; }

        public double Costo { get; set; }

        public double Total { get; set; }

        public ElementoTabla(int cantidad)
        {
            Codigo = $"Codigo-{cantidad}";
            
            Cantidad = cantidad + 1;

            var rand = new Random();
            Costo = rand.Next(cantidad + 100);

            Total = (Cantidad * Costo);
        }
    }
}
