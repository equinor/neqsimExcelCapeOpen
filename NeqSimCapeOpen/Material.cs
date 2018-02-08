using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenThermo
{

    // test
    class Material : ICapeThermoMaterial
    {
        Material mat;
        
        public void ClearAllProps()
        {
        }

        public object CreateMaterial()
        {
            mat = new Material();
            return mat;
        }

        public void CopyFromMaterial(ref object matIn)
        {
            //mat.

        }

        public void GetOverallProp(String property, String basis, ref object vals)
        {
        }

        public void GetOverallTPFraction(ref double temperature, ref double pressure, ref object composition)
        {
        }

        public void GetPresentPhases(ref object phaseLabels, ref object composition)
        {
        }


        public void GetSinglePhaseProp(string property, string phaseLabel, string basis, ref object val)
        {
        }

        public void SetSinglePhaseProp(string property, string phaseLabel, string basis, object val)
        {
        }


        public void GetTPFraction(string phaseLabel, ref double temperature, ref double pressure, ref object composition)
        {
        }


        public void GetTwoPhaseProp(string property, object paseLabels, string basis, ref object results)
        {
        }


        public void SetOverallProp(string property, string basis, object val)
        {
        }


        public void SetPresentPhases(object phaseLabels, object phaseStatus)
        {
        }


        public void SetSinglePhaseProp(string property, object paseLabels, string basis, ref object results)
        {
        }

        public void SetTwoPhaseProp(string property, object paseLabels, string basis,  object results)
        {
        }
        
    }
}
