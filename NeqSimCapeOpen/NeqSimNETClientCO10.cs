using System;
using NeqSimNET;
using CAPEOPEN100;

namespace CapeOpenThermo
{
    public class NeqSimNETClientCO10 : ICapeIdentification, ICapeThermoPropertyPackage

    {
        public NeqSimNETService neqsimService = null;
        public string[] constPropList, compNames, formulae, compIds, CAS;

        string[] phaseL = { "Vapor", "Liquid", "Liquid2" };
        string[] stateOfA = { "Vapor", "Liquid", "Liquid" };
        string[] keyCompID = { "", "", "Water" };

        string[] densityDescr = { "", "Light", "Heavy" };
        public string[] excludedIDs = { "", "Water", "" };

        public int[] oilFractionIDs = null;

        public string[] properties = null;
        public string[] twoProp = null;
        public int nNumComp = 2;
        public int numPhases = 3;
        public string componentDescription, componentName;

        int initNumb = 0, oldInitNumb = 0;
        public NeqSimNETClientCO10()
        {
        }

        public NeqSimNETClientCO10(String package)
        {
            neqsimService = new NeqSimNETService();
            neqsimService.readFluidFromGQIT(Convert.ToInt32(package));
            neqsimService.setPackageID(Convert.ToInt32(package));

            properties = neqsimService.readCapeOpenProperties10();

            nNumComp = neqsimService.getNumberOfComponents();
            twoProp = null;// new string[] { "surfaceTension" };//, "kvalue", "logKvalue" };
            oilFractionIDs = neqsimService.getOilFractionIDs();
            constPropList = new String[] { "liquidDensityAt25C", "molecularWeight", "normalBoilingPoint", "acentricFactor", "criticalPressure", "criticalTemperature", "criticalVolume" };
          
        }


        public object ValidityCheck(object props1, object props2)
        {
            return null;
        }

        public object PropCheck(object props1, object props2)
        {
            return null;
        }

        public object GetUniversalConstant(object props1, object props2)
        {
            return null;
        }
        public object GetPropList()
        {
            return properties;
        }
        public object GetPhaseList()
        {
            return phaseL;
        }

      
        public void GetComponentList(ref object compIds, ref object formulae, ref object names, ref object boilTemps, ref object molwts, ref object casnos)
         {
                 neqsimService.GetCompoundList(ref compIds, ref formulae, ref names, ref boilTemps, ref molwts, ref casnos);
         }

        public  void CalcEquilibrium(object specification1, string specification2, object name)
        {
        }

        public void CalcProp(object material, object props, object phaseLabels, string name2)
        {
            try
            {
              
                string[] lables = (string[])phaseLabels;
                for (int i = 0; i < lables.Length; i++)
                {
                    if (lables[i].Equals("Overall"))
                    {
                        CalcSinglePhaseProp((ICapeThermoMaterialObject)material, props, lables[i]);
                    }
                    CalcSinglePhaseProp((ICapeThermoMaterialObject)material, props, lables[i]);
                }
            }
            catch (Exception e)
            {
                String w = e.Message;
                throw e;
            }
           
        }

        
        public void CalcSinglePhaseProp(ICapeThermoMaterialObject material, object props, string phaseLabel)
        {
            try
            {
                string[] tempString = (string[])props;
                int length = tempString.Length;
                double temperature = 0;
                double pressure = 0;
                object composition = null;
                try
                {
                    temperature = material.GetProp("temperature", "Overall", null, "Mixture", "")[0]; //"ref composition);
                    pressure = material.GetProp("pressure", "Overall", null, "Mixture", "")[0];
                    composition = material.GetProp("fraction", phaseLabel, null, "Mixture", "mole");
                }
                catch (Exception e)
                {
                    String w = e.Message;
                    throw e;
                }

                if (pressure > 2e8)
                {
                    pressure = 2e8;
                    return;
                }


                Boolean doInit = true;

                for (int i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("logFugacityCoefficient") || tempString[i].Equals("fugacityCoefficient") || tempString[i].StartsWith("molecularWeight") || tempString[i].Equals("volume.Dpressure") || tempString[i].Equals("density") || tempString[i].Equals("compressibilityFactor") || tempString[i].Equals("density.Dpressure") || tempString[i].Equals("density.Dmoles") || tempString[i].Equals("volume") || tempString[i].Equals("helmholtzEnergy") || tempString[i].Equals("gibbsEnergy") || tempString[i].Equals("viscosity") || tempString[i].Equals("thermalConductivity")) initNumb = 1;
                    else if (tempString[i].Equals("logFugacityCoefficient.Dmoles")) initNumb = 3;
                    else initNumb = 2;

                    if (initNumb > oldInitNumb || (neqsimService.checkIfInitNeed(temperature, pressure / 1.0e5, (double[])composition, phaseLabel)))
                    {
                        oldInitNumb = initNumb;
                        neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[])composition);
                        neqsimService.init(phaseLabel, initNumb);
                        doInit = false;
                    }
                    else
                    {
                        doInit = false;
                    }
                }

                for (int i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("fugacityCoefficient"))
                    {
                        double[] fugCoefs = neqsimService.getFugacityCoefficients(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefs);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient"))
                    {
                        //double[] lnPhiTemp = new double[nNumComp];
                        double[] lnPhiTemp = neqsimService.getLogFugacityCoefficients(phaseLabel, doInit);
                        //for (int k = 0; k < fugCoef.Length; k++)
                        //{
                        //    lnPhiTemp[k] = Math.Log(fugCoef[k]);
                        // }
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", lnPhiTemp);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dtemperature"))
                    {
                        double[] fugCoefsdT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, doInit);

                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefsdT);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dpressure"))
                    {
                        double[] fugCoefsdP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefsdP);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dmoles"))
                    {
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, doInit));
                        continue;
                    }
                    else if (tempString[i].Equals("density"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("compressibilityFactor"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getCompressibilityFactor(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dtemperature"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensitydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dpressure"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensitydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight"))
                    {
                        double[] molWtrer = new double[1];
                        molWtrer[0] = neqsimService.getMolecularWeight(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", molWtrer);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight.Dtemperature") || tempString[i].Equals("molecularWeight.Dpressure"))
                    {
                        double[] molWtrer = new double[1];
                        molWtrer[0] = 0.0;
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", molWtrer);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight.Dmoles"))
                    {
                        double[] Mdn = neqsimService.getMolecularWeightdN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", Mdn);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getDensitydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                        continue;
                    }
                    else if (tempString[i].Equals("volume"))
                    {
                        double[] volume = new double[1];
                        volume[0] = neqsimService.getVolume(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volume);
                        continue;
                    }
                    if (tempString[i].Equals("volume.Dtemperature"))
                    {
                        double[] volumedT = new double[1];
                        volumedT[0] = neqsimService.getVolumedT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volumedT);
                        continue;
                    }
                    else if (tempString[i].Equals("volume.Dpressure"))
                    {
                        double[] volumedP = new double[1];
                        volumedP[0] = neqsimService.getVolumedP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volumedP);
                        continue;
                    }
                    else if (tempString[i].Equals("heatCapacity"))
                    {
                        double[] heatCapacityCp = new double[1];
                        heatCapacityCp[0] = neqsimService.getHeatCapacityCp(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", heatCapacityCp);
                        continue;
                    }
                    else if (tempString[i].Equals("heatCapacityCv"))
                    {
                        double[] heatCapacityCv = new double[1];
                        heatCapacityCv[0] = neqsimService.getHeatCapacityCv(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", heatCapacityCv);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy.Dtemperature"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy.Dpressure"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getEnthalpydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dpressure"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dtemperature"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getEntropydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                        continue;
                    }
                    else if (tempString[i].Equals("helmholtzEnergy"))
                    {
                        double[] H = new double[1];
                        H[0] = neqsimService.getHelmholtzEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", H);
                        continue;
                    }
                    else if (tempString[i].Equals("jouleThomsonCoefficient"))
                    {
                        double[] JT = new double[1];
                        JT[0] = neqsimService.getJouleThomsonCoefficient(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", JT);
                        continue;
                    }
                    else if (tempString[i].Equals("gibbsEnergy"))
                    {
                        double[] gibbsEnergy = new double[1];
                        gibbsEnergy[0] = neqsimService.getGibbsEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", gibbsEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("energy"))
                    {
                        double[] internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", internalEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("energy.Dtemperature"))
                    {
                        double[] internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", internalEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("speedOfSound"))
                    {
                        double[] speedOfSound = new double[1];
                        speedOfSound[0] = neqsimService.getSpeedOfSound(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", speedOfSound);
                        continue;
                    }

                    else if (tempString[i].Equals("viscosity"))
                    {
                        double[] viscosity = new double[1];
                        viscosity[0] = neqsimService.getViscosity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", viscosity);
                        continue;
                    }
                    else if (tempString[i].Equals("thermalConductivity"))
                    {
                        double[] conductivity = new double[1];
                        conductivity[0] = neqsimService.getThermalConductivity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", conductivity);
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                String w = e.Message;
                throw e;
            }
        }

        public virtual int GetNumCompounds()
        {
            return nNumComp;
        }

        public Object GetComponentConstant(object material, object props)
        {
            int numberOfComponents = 0;

            numberOfComponents = GetNumCompounds();

            string[] propsName = (string[])props;
            string[] tempString = (string[])props;
            int propLength = tempString.Length;
            if (propLength == 0) return null;

            Object[] propsVal = new Object[numberOfComponents];

            for (int i = 0; i < propLength; i++)
            {
                for (int j = 0; j < numberOfComponents; j++)
                {
                    if (propsName[i].Equals("criticalVolume")) propsVal[j] = neqsimService.getCriticalVolumes()[j];
                    if (propsName[i].Equals("criticalTemperature")) propsVal[j] = neqsimService.getCriticalTemperatures()[j];
                    if (propsName[i].Equals("criticalPressure")) propsVal[j] = neqsimService.getCriticalPressures()[j];
                    if (propsName[i].Equals("acentricFactor")) propsVal[j] = neqsimService.getAcentricFactor()[j];
                    if (propsName[i].Equals("normalBoilingPoint")) propsVal[j] = neqsimService.getNormalBoilingPoint()[j];
                    if (propsName[i].Equals("molecularWeight")) propsVal[j] = neqsimService.getMolecularWeight()[j];
                    if (propsName[i].Equals("liquidDensityAt25C")) propsVal[j] = neqsimService.getLiquidDensityAt25C()[j];
                }
            }
            return propsVal;
        }

        public Object GetComponentConstant2(object props, object compIds)
        {
            int numberOfComponents = 0;

            if (compIds == null) numberOfComponents = GetNumCompounds();
            else
            {
                numberOfComponents = ((string[])compIds).Length;
            }

            string[] propsName = (string[])props;
            string[] tempString = (string[])props;
            int propLength = tempString.Length;
            if (propLength == 0) return null;

            Object[] propsVal = new Object[numberOfComponents];

            for (int i = 0; i < propLength; i++)
            {

                for (int j = 0; j < numberOfComponents; j++)
                {
                    if (propsName[i].Equals("criticalVolume")) propsVal[j] = neqsimService.getCriticalVolumes()[j];
                    if (propsName[i].Equals("criticalTemperature")) propsVal[j] = neqsimService.getCriticalTemperatures()[j];
                    if (propsName[i].Equals("criticalPressure")) propsVal[j] = neqsimService.getCriticalPressures()[j];
                    if (propsName[i].Equals("acentricFactor")) propsVal[j] = neqsimService.getAcentricFactor()[j];
                    if (propsName[i].Equals("normalBoilingPoint")) propsVal[j] = neqsimService.getNormalBoilingPoint()[j];
                    if (propsName[i].Equals("molecularWeight")) propsVal[j] = neqsimService.getMolecularWeight()[j];
                    if (propsName[i].Equals("liquidDensityAt25C")) propsVal[j] = neqsimService.getLiquidDensityAt25C()[j];
                }
            }
            return propsVal;
        }

        public string ComponentName
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }
        public String ComponentDescription
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }
    }
}
