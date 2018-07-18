using System.Collections.Generic;
using System.Linq;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class PortCollectionBaseClass : ICapeCollection, ICapeIdentification
    {
        // PortBaseClass port1, port2, port3;
        private readonly List<PortBaseClass> ports = new List<PortBaseClass>();

        public object Item(object item)
        {
            if (item.ToString().Equals("1")) return ports.ElementAt(0);
            if (item.ToString().Equals("2")) return ports.ElementAt(1);
            if (item.ToString().Equals("3")) return ports.ElementAt(2);
            if (item.ToString().Equals("4")) return ports.ElementAt(3);
            for (var i = 0; i < ports.Count; i++)
                if (item.ToString().Equals(ports.ElementAt(i).ComponentName))
                    return ports.ElementAt(i);
            //if (item.ToString().Equals("Port1")) return port1;
            //if (item.ToString().Equals("Port2")) return port2;
            //if (item.ToString().Equals("Port3")) return port3;
            return ports.ElementAt(0);
        }


        public int Count()
        {
            return ports.Count;
        }

        public string ComponentDescription { get; set; } = "Portlist";


        public string ComponentName { get; set; } = "Portlist2";

        public void addPort(string portName, CapePortDirection direction)
        {
            ports.Add(new PortBaseClass(portName, direction));
        }
    }
}