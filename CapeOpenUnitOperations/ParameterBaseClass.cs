using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterBaseClass : ICapeParameter, ICapeIdentification
    {
        private readonly ICapeParameterSpec paramSpec = new ParameterBaseSpec();

        public ParameterBaseClass(string inpName, string inpcomponentDescription, CapeParamMode inpParamMode,
            ICapeParameterSpec inpparamSpec, object inpparamValue)
        {
            ComponentName = inpName;
            ComponentDescription = inpcomponentDescription;
            Mode = inpParamMode;
            paramSpec = inpparamSpec;
            value = inpparamValue;
        }

        public string ComponentDescription { get; set; } = "Pressure in unit";


        public string ComponentName { get; set; } = "Pressure";

        public object Specification => paramSpec;

        public CapeValidationStatus ValStatus { get; } = CapeValidationStatus.CAPE_VALID;

        public bool Validate(ref string errText)
        {
            //  return false;
            return true;
        }

        public CapeParamMode Mode { get; set; } = CapeParamMode.CAPE_INPUT;

        public void Reset()
        {
            value = null;
        }


        public object value { get; set; }
    }
}