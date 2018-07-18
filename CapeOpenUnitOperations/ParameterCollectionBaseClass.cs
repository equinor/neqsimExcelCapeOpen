using System.Collections.Generic;
using System.Linq;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterCollectionBaseClass : ICapeCollection, ICapeIdentification
    {
        // ICapeParameter param1 = new ParameterBaseClass("NeqSimUnitID", "Online unit ID from NeqSim", CapeParamMode.CAPE_INPUT, new ParameterIntSpec(),1);
        private readonly List<ICapeParameter> parameters = new List<ICapeParameter>();

        public object Item(object item)
        {
            if (item.ToString().Equals("1")) return parameters.ElementAt(0);
            if (item.ToString().Equals("2")) return parameters.ElementAt(1);
            return parameters.ElementAt(0);
        }

        public int Count()
        {
            return parameters.Count;
        }

        public string ComponentDescription { get; set; } = "Parameterlist";


        public string ComponentName { get; set; } = "Parameterlist2";

        public void AddParameter(string inpName, string inpcomponentDescription, CapeParamMode inpParamMode,
            ICapeParameterSpec inpparamSpec, object inpparamValue)
        {
            parameters.Add(new ParameterBaseClass(inpName, inpcomponentDescription, inpParamMode, inpparamSpec,
                inpparamValue));
        }
    }
}