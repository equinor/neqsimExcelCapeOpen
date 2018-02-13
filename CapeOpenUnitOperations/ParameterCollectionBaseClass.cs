using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterCollectionBaseClass : ICapeCollection, ICapeIdentification
    {

        string componentDescription = "Parameterlist";
        string componentName = "Parameterlist2";
       // ICapeParameter param1 = new ParameterBaseClass("NeqSimUnitID", "Online unit ID from NeqSim", CapeParamMode.CAPE_INPUT, new ParameterIntSpec(),1);
        List<ICapeParameter> parameters = new List<ICapeParameter>();

        public Object Item(object item)
        {
            if (item.ToString().Equals("1")) return parameters.ElementAt<ICapeParameter>(0);
            if (item.ToString().Equals("2")) return parameters.ElementAt<ICapeParameter>(1);
            else return parameters.ElementAt<ICapeParameter>(0);
        }

        public void AddParameter(string inpName, string inpcomponentDescription, CapeParamMode inpParamMode, ICapeParameterSpec inpparamSpec, object inpparamValue)
        {
            parameters.Add(new ParameterBaseClass(inpName, inpcomponentDescription, inpParamMode, inpparamSpec, inpparamValue));
        }

        public int Count()
        {
            return parameters.Count;
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
