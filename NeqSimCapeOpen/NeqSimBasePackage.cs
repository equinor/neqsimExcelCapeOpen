using System;
using CAPEOPEN110;
using NeqSimNET;

namespace CapeOpenThermo
{
    public class NeqSimBasePackage : ICapeIdentification, ICapeThermoCompounds, ICapeThermoPhases,
        ICapeThermoMaterialContext, ICapeThermoEquilibriumRoutine, ICapeThermoPropertyRoutine, ICapeThermoUniversalConstant
    {
        private readonly double[] boilT = null;
        private readonly double[] molw = null;
        public string componentDescription, componentName;
        public string[] constPropList, compNames, formulae, compIds, CAS;

        private string[] densityDescr = {"", "Light", "Heavy"};
        public string[] excludedIDs = {"", "Water", ""};
        private readonly string[] keyCompID = {"", "", "Water"};
        public ICapeThermoMaterial material;

        public NeqSimNETService neqsimService = null;
        public int nNumComp = 2;
        public int numPhases = 3;

        public int[] oilFractionIDs = null;

        private readonly string[] phaseL = {"Vapor", "Liquid", "Liquid2"};

        public string[] properties = null;
        private readonly string[] stateOfA = {"Vapor", "Liquid", "Liquid"};
        public string[] twoProp = null;

        public string ComponentDescription
        {
            get => componentDescription;
            set => componentDescription = value;
        }

        public string ComponentName
        {
            get => componentName;
            set => componentName = value;
        }

        public virtual void GetCompoundList(ref object compIds, ref object formulae, ref object names,
            ref object boilTemps, ref object molwts, ref object casnos)
        {
            names = compNames;
            formulae = this.formulae;
            compIds = this.compIds;
            boilTemps = boilT;
            molwts = molw;
            casnos = CAS;
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

        public object GetUniversalConstantList()
        {
            return null;
        }

        public object GetUniversalConstant(string constName)
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


        public object GetConstPropList()
        {
            return constPropList;
        }

        public virtual int GetNumCompounds()
        {
            return nNumComp;
        }

        public virtual void CalcEquilibrium(object specification1, object specification2, string name)
        {
        }

        public virtual bool CheckEquilibriumSpec(object specification1, object specification2, string name)
        {
            return false;
        }

        public void UnsetMaterial()
        {
            material = null;
        }

        public void SetMaterial(object inMat)
        {
            // material = inMat;// new TestMaterial();
            material = (ICapeThermoMaterial) inMat;
        }

        public int GetNumPhases()
        {
            return numPhases;
        }


        public void GetPhaseList(ref object phaseLabels, ref object stateOfAggregation, ref object keyCompoundId)
        {
            phaseLabels = phaseL;
            stateOfAggregation = stateOfA;
            keyCompoundId = keyCompID;
        }


        public object GetPhaseInfo(string phaseLabel, string phaseAttribute)
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

        public virtual void CalcAndGetLnPhi(string phaseLabel, double temperature, double pressure, object moleNumbers,
            int fFlags, ref object lnPhi, ref object lnPhiDT, ref object lnPhiDP, ref object lnPhiDn)
        {
            var lnPhiTemp = new double[2];
            var lnPhiDTTemp = new double[2];
            var lnPhiDPTemp = new double[2];
            var lnPhiDnTemp = new double[2, 2];

            var fugCoef = new double[2];
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
            if (fFlags == 1) lnPhi = lnPhiTemp;
            else if (fFlags == 2) lnPhiDT = lnPhiDTTemp;
            else if (fFlags == 4) lnPhiDP = lnPhiDPTemp;
            else if (fFlags == 8) lnPhiDn = lnPhiDnTemp;
        }

        public virtual void CalcSinglePhaseProp(object props, string phaseLabel)
        {
            try
            {
                var tempString = (string[]) props;
                var length = tempString.Length;
                double temperature = 0;
                double pressure = 0;
                object composition = null;
                material.GetTPFraction(phaseLabel, ref temperature, ref pressure, ref composition);

                var enthalpy = new double[1];
                enthalpy[0] = 1.1;

                var fugCoef = new double[2];
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

                var lnPhiTemp = new double[2];
                lnPhiTemp[0] = Math.Log(fugCoef[0]);
                lnPhiTemp[1] = Math.Log(fugCoef[1]);
                var lnPhiDTTemp = new double[2];
                var lnPhiDPTemp = new double[2];
                var lnPhiDnTemp = new double[2, 2];
                for (var i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("enthalpy"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    if (tempString[i].Equals("entropy"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    if (tempString[i].Equals("fugacityCoefficient"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoef);
                    if (tempString[i].Equals("logFugacityCoefficient"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dtemperature"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDTTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dpressure"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDPTemp);
                    if (tempString[i].Equals("logFugacityCoefficient.Dmoles"))
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiDnTemp);
                    //else material.SetSinglePhaseProp(tempString[i], phaseLabel, "UNDEFINED", and);
                }
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }
        }

        public virtual void CalcTwoPhaseProp(object props, object phaseLabels)
        {
            var and = new object[2];
            and[0] = 4.0;
            and[1] = 0.1;


            double temperature = 0;
            double pressure = 0;
            object composition = null;
            material.GetTPFraction("Vapor", ref temperature, ref pressure, ref composition);
            //  material.SetTwoPhaseProp("kvalue", ((string[]) phaseLabels)[0], "UNDEFINED", and);
            material.SetTwoPhaseProp("kvalue", phaseLabels, "", and);
            // material.SetTwoPhaseProp("kvalue", "Liquid", "UNDEFINED", and);
        }

        public bool CheckSinglePhasePropSpec(string property, string phaseLabel)
        {
            // char[] charsToTrim = {'\''};
            // string NewString = property.TrimEnd(charsToTrim);
            // if (neqsimService.CapeOpenProperties.Contains(property)) return true;
            // else return true;
            return true;
        }

        public bool CheckTwoPhasePropSpec(string property, object phaseLabel)
        {
            if (property.Equals("kvalue") && phaseLabel.Equals("Vapor")) return true;
            if (property.Equals("kvalue") && phaseLabel.Equals("Liquid")) return true;
            return false;
        }

        public object GetSinglePhasePropList()
        {
            return properties;
        }


        public object GetTwoPhasePropList()
        {
            return twoProp;
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