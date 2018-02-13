using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class PortBaseClass : ICapeUnitPort, ICapeIdentification
    {
        string componentDescription = "Port def";
        string componentName;// = "Port1";
        public ICapeThermoMaterial material = null;

        CapePortDirection directionStream = CapePortDirection.CAPE_INLET;

        public PortBaseClass(string name, CapePortDirection directionIn)
        {
            componentName = name;
            directionStream = directionIn;
        }

        public CapePortType portType
        {
            get
            {
                return CapePortType.CAPE_MATERIAL;
           }

        }

        public CapePortDirection direction
        {
            get
            {
                return directionStream;
            }

        }

        public Object connectedObject
        {
            get
            {
                return material;
            }
        }

        public void Connect(Object connectedObject)
        {
            material = (ICapeThermoMaterial) connectedObject;
        }

        public void Disconnect()
        {
            material = null;
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
