using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterBaseClass : ICapeParameter, ICapeIdentification
    {
        string componentDescription = "Pressure in unit";
        string componentName = "Pressure";

        object paramValue = null;
        CapeParamMode paramMode = CapeParamMode.CAPE_INPUT;
        ICapeParameterSpec paramSpec = new ParameterBaseSpec();
        CAPEOPEN110.CapeValidationStatus ValStatus2 = CAPEOPEN110.CapeValidationStatus.CAPE_VALID;

        public ParameterBaseClass(string inpName, string inpcomponentDescription, CapeParamMode inpParamMode, ICapeParameterSpec inpparamSpec, object inpparamValue)
        {
            this.componentName = inpName;
            this.componentDescription = inpcomponentDescription;
            this.paramMode = inpParamMode;
            this.paramSpec = inpparamSpec;
            this.paramValue = inpparamValue;
        }

        public object Specification
        {
            get
            {
                return paramSpec;
            }
          
        }

        public CapeValidationStatus ValStatus
        {
            get
            {
                return ValStatus2;
                // throw new NotImplementedException();
            }
        }

        public bool Validate(ref String errText)
        {
          //  return false;
           return true;
        }

        public CapeParamMode Mode
        {
            get
            {
                return paramMode;
                // throw new NotImplementedException();
            }
              set
            {
                paramMode = value;
                // throw new NotImplementedException();
            }
        }

        public void Reset()
        {
            paramValue = null;
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


        public Object value
        {
            get
            {
                return paramValue;
                // throw new NotImplementedException();
            }
            set
            {
                paramValue = value;
                //throw new NotImplementedException();
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
