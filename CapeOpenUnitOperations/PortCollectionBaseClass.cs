using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class PortCollectionBaseClass: ICapeCollection, ICapeIdentification
    {
        string componentDescription = "Portlist";
        string componentName = "Portlist2";
       // PortBaseClass port1, port2, port3;
        List<PortBaseClass> ports = new List<PortBaseClass>();

        public PortCollectionBaseClass()
        {
       //     port1 = new PortBaseClass("Port1", CapePortDirection.CAPE_INLET);
       //     port2 = new PortBaseClass("Port2", CapePortDirection.CAPE_OUTLET);
       //     port3 = new PortBaseClass("Port3", CapePortDirection.CAPE_INLET);
        }

        public void addPort(String portName, CapePortDirection direction)
        {
            ports.Add(new PortBaseClass(portName, direction));
        }

        public Object Item(object item)
        {
            if (item.ToString().Equals("1")) return ports.ElementAt<PortBaseClass>(0);
            if (item.ToString().Equals("2")) return ports.ElementAt<PortBaseClass>(1);
            if (item.ToString().Equals("3")) return ports.ElementAt<PortBaseClass>(2);
            if (item.ToString().Equals("4")) return ports.ElementAt<PortBaseClass>(3);
            for (int i = 0; i < ports.Count; i++)
            {
                if(item.ToString().Equals(ports.ElementAt<PortBaseClass>(i).ComponentName)) return ports.ElementAt<PortBaseClass>(i);
            }
            //if (item.ToString().Equals("Port1")) return port1;
            //if (item.ToString().Equals("Port2")) return port2;
            //if (item.ToString().Equals("Port3")) return port3;
            return ports.ElementAt<PortBaseClass>(0);
        }


        public int Count()
        {
            return ports.Count;
        }

        public String ComponentDescription
        {
            get
            {
                return componentDescription;
                // throw new NotImplementedException();
            }
            set
            {
                componentDescription = value;
                //  throw new NotImplementedException();
            }
        }



        public String ComponentName
        {
            get
            {
                return componentName;
                // throw new NotImplementedException();
            }
            set
            {
                componentName = value;
                //throw new NotImplementedException();
            }
        }
    }
}
