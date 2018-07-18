using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class PortBaseClass : ICapeUnitPort, ICapeIdentification
    {
        public ICapeThermoMaterial material;

        public PortBaseClass(string name, CapePortDirection directionIn)
        {
            ComponentName = name;
            direction = directionIn;
        }

        public string ComponentDescription { get; set; } = "Port def";


        public string ComponentName { get; set; }

        public CapePortType portType => CapePortType.CAPE_MATERIAL;

        public CapePortDirection direction { get; } = CapePortDirection.CAPE_INLET;

        public object connectedObject => material;

        public void Connect(object connectedObject)
        {
            material = (ICapeThermoMaterial) connectedObject;
        }

        public void Disconnect()
        {
            material = null;
        }
    }
}