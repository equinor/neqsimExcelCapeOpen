using System;
using CAPEOPEN110;
using NeqSimNET;


namespace CapeOpenThermo
{
    public class NeqSimBasePackage : ICapeIdentification, ICapeThermoCompounds, ICapeThermoPhases, ICapeThermoMaterialContext, ICapeThermoEquilibriumRoutine, ICapeThermoPropertyRoutine
    {

        public NeqSimNETService neqsimService = null;
        public string[] constPropList, compNames, formulae, compIds, CAS ;
        double[] boilT = null, molw = null;

        string[] phaseL = { "Vapor", "Liquid", "Liquid2" };
        string[] stateOfA = { "Vapor", "Liquid", "Liquid" };
        string[] keyCompID = { "", "", "Water"};

        string[] densityDescr = { "", "Light", "Heavy" };
        public string[] excludedIDs = { "", "Water", "" };

        public int[] oilFractionIDs = null;
        public ICapeThermoMaterial material = null;
        
        public string[] properties = null;
        public string[] twoProp = null;
        public int nNumComp = 2;
        public int numPhases = 3;
        public string componentDescription, componentName;

        public NeqSimBasePackage(){
        }

        public int GetNumPhases()
        {
            return numPhases;
        }

        public virtual void CalcEquilibrium(object specification1, object specification2, string name)
        {
        }

        public virtual Boolean CheckEquilibriumSpec(object specification1, object specification2, string name)
        {
            return false;
        }


        public void GetPhaseList(ref object phaseLabels, ref object stateOfAggregation, ref object keyCompoundId){
            phaseLabels = phaseL;
            stateOfAggregation = stateOfA;
            keyCompoundId = keyCompID;
        }


        public Object GetPhaseInfo(string phaseLabel, string phaseAttribute)
        {
            if (phaseLabel.Equals("Vapor") && phaseAttribute.Equals("StateOfAggregation")) return "Vapor";
            if (phaseLabel.Equals("Liquid") && phaseAttribute.Equals("StateOfAggregation")) return "Liquid";
            if (phaseLabel.Equals("Liquid2") && phaseAttribute.Equals("StateOfAggregation")) return "Liquid";

            if (phaseLabel.Equals("Vapor") && phaseAttribute.Equals("KeyCompoundId")) return null;
            if (phaseLabel.Equals("Liquid") && phaseAttribute.Equals("KeyCompoundId")) return null;
            if (phaseLabel.Equals("Liquid2") && phaseAttribute.Equals("KeyCompoundId")) return "Water";

            if (phaseLabel.Equals("Vapor") && phaseAttribute.Equals("DensityDescription")) return "";
            if (phaseLabel.Equals("Liquid") && phaseAttribute.Equals("DensityDescription")) return "Light";
            if (phaseLabel.Equals("Liquid2") && phaseAttribute.Equals("DensityDescription")) return "Heavy";

            if (phaseLabel.Equals("Vapor") && phaseAttribute.Equals("UserDescription")) return "Vapour Phase";
            if (phaseLabel.Equals("Liquid") && phaseAttribute.Equals("UserDescription")) return "Liquid Phase";
            if (phaseLabel.Equals("Liquid2") && phaseAttribute.Equals("UserDescription")) return "Aqueous Phase";
            
            if (phaseLabel.Equals("Vapor") && phaseAttribute.Equals("ExcludedCompoundId")) return null;
            if (phaseLabel.Equals("Liquid") && phaseAttribute.Equals("ExcludedCompoundId")) return "Water";
            if (phaseLabel.Equals("Liquid2") && phaseAttribute.Equals("ExcludedCompoundId")) return null;

            return null;
        }

        public virtual void GetCompoundList(ref object compIds, ref object formulae, ref object names, ref object boilTemps, ref object molwts, ref object casnos)
        {
            names = this.compNames;
            formulae = this.formulae;
            compIds = this.compIds;
            boilTemps = this.boilT;
            molwts = this.molw;
            casnos = this.CAS;
        }

        public virtual void CalcAndGetLnPhi(string phaseLabel, double temperature, double pressure, object moleNumbers, int fFlags, ref object lnPhi, ref object lnPhiDT, ref object lnPhiDP, ref object lnPhiDn)
        {
            double[] lnPhiTemp= new double[2];
            double[] lnPhiDTTemp= new double[2];
            double[] lnPhiDPTemp= new double[2];
            double[,] lnPhiDnTemp = new double[2,2];

            double[] fugCoef = new double[2];
            if (phaseLabel.Equals("Vapor"))
            {
                fugCoef[0] = 1.0;
                fugCoef[1] = 1.0;
            }
            else
            {
                fugCoef[0] = 10.1;
                fugCoef[1] = 0.06;
            }

            lnPhiTemp[0] = Math.Log(fugCoef[0]);
            lnPhiTemp[1] = Math.Log(fugCoef[1]);

            if (fFlags == 0) return;
            else if (fFlags == 1) lnPhi = lnPhiTemp;
            else if (fFlags == 2) lnPhiDT = lnPhiDTTemp;
            else if (fFlags == 4) lnPhiDP = lnPhiDPTemp;
            else if (fFlags == 8) lnPhiDn = lnPhiDnTemp;
            return;
        }

        public virtual void CalcSinglePhaseProp(object props, string phaseLabel)
        {
            try
            {
                string[] tempString = (string[])props;
                int length = tempString.Length;
                double temperature = 0;
                double pressure = 0;
                object composition = null;
                material.GetTPFraction(phaseLabel, ref temperature, ref pressure, ref composition);

                double[] enthalpy = new double[1];
                enthalpy[0] = 1.1;

                double[] fugCoef = new double[2];
                if (phaseLabel.Equals("Vapor"))
                {
                    fugCoef[0] = 1.0;
                    fugCoef[1] = 1.0;
                }
                else
                {
                    fugCoef[0] = 10.1;
                    fugCoef[1] = 0.06;
                }
                double[] lnPhiTemp = new double[2];
                lnPhiTemp[0] = Math.Log(fugCoef[0]);
                lnPhiTemp[1] = Math.Log(fugCoef[1]);
                double[] lnPhiDTTemp = new double[2];
                double[] lnPhiDPTemp = new double[2];
                double[,] lnPhiDnTemp = new double[2, 2];
                for (int i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("enthalpy")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    if (tempString[i].Equals("entropy")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    if (tempString[i].Equals("fugacityCoefficient")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoef);
                    if (tempString[i].Equals("logFugacityCoefficient")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dtemperature")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDTTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dpressure")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDPTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dmoles")) material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDnTemp);
                    //else material.SetSinglePhaseProp(tempString[i], phaseLabel, "UNDEFINED", and);
                }
            }
            catch (Exception e)
            {
               String w = e.Message;
               throw e;
            }

        }

        virtual public void CalcTwoPhaseProp(object props, object phaseLabels)
        {
            object[] and = new object[2];
            and[0] = 4.0;
            and[1] = 0.1;


            double temperature=0;
            double pressure=0;
            object composition=null;
            material.GetTPFraction("Vapor",   ref temperature, ref pressure, ref composition);
          //  material.SetTwoPhaseProp("kvalue", ((string[]) phaseLabels)[0], "UNDEFINED", and);
            material.SetTwoPhaseProp("kvalue", phaseLabels, "", and);
           // material.SetTwoPhaseProp("kvalue", "Liquid", "UNDEFINED", and);
        }

        public Boolean CheckSinglePhasePropSpec(string property, string phaseLabel)
        {
           // char[] charsToTrim = {'\''};
           // string NewString = property.TrimEnd(charsToTrim);
           // if (neqsimService.CapeOpenProperties.Contains(property)) return true;
           // else return true;
          return true;

        }

        public Boolean CheckTwoPhasePropSpec(string property, object phaseLabel)
        {
            if (property.Equals("kvalue") && phaseLabel.Equals("Vapor")) return true;
            if (property.Equals("kvalue") && phaseLabel.Equals("Liquid")) return true;
            else return false;
        }

        public Object GetSinglePhasePropList()
        {
            return properties;
        }


        public object GetTwoPhasePropList()
        {
            return twoProp;
        }
        public void GetPDependentProperty(object compIds, double formulae, object names, ref object outs)
        {
        }

        public object GetPDependentPropList()
        {
            return null;
        }

        public object GetTDependentPropList()
        {
            return null;
        }

        public void GetTDependentProperty(object compIds, double formulae, object names, ref object outs)
        {
        }

        public virtual object GetCompoundConstant(object props, object compIds)
        {
            return null;
        }

        

        public Object GetConstPropList()
        {
            return constPropList;
        }

        public virtual int GetNumCompounds()
        {
            return nNumComp;
        }

        public string ComponentDescription
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

        public string ComponentName
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

        public void UnsetMaterial()
        {
            material = null;
        }

        public void SetMaterial(Object  inMat)
        {
            // material = inMat;// new TestMaterial();
            material = (ICapeThermoMaterial)inMat;
        }
        /*
        public void SetPetroProp(string propertyID, object compID, string basis, object values)
        {
            //if (propertyID[0].Equals("LiquidDensity"));
            return;
        }

        public object GetPetroProp(string propertyID, object compID, string basis)
        {
            string[] compIDNames = (string[])compID;
            double[] propsVal = new double[compIDNames.Length];

             for (int i = 0; i < propsVal.Length; i++)
             {
                if (propertyID.Equals("LiquidDensity")) propsVal[0]=760.0;
                if (propertyID.Equals("MolecularWeight")) propsVal[0] = 120.0;
                if (propertyID.Equals("NormalBoilingPoint")) propsVal[0] = 220.0;
             }
             return propsVal;
        }

        public void Characterize()
        {

        }

        public void RemovePetroProp(string propertyID, object compID)
        {

        }

        public void DefineFromPetroFractions(object orginFractionsSet)
        {

        }
         * 
         * */

        public object GetUnitType()
        {
            return null;

        }

        /*
        public object GetUniversalConstantList()
        {
            string[] temp = null;
            return temp;
        }

        public object GetUniversalConstant(String name)
        {
            object temp = null;
            return temp;
        }*/
    }

   
}
    
