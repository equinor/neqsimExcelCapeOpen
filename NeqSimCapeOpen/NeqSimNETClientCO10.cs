using System;
using CAPEOPEN100;
using NeqSimNET;

namespace CapeOpenThermo
{
    public class NeqSimNETClientCO10 : ICapeIdentification, ICapeThermoPropertyPackage

    {
        public string componentDescription, componentName;
        public string[] constPropList, compNames, formulae, compIds, CAS;

        private string[] densityDescr = {"", "Light", "Heavy"};
        public string[] excludedIDs = {"", "Water", ""};

        private int initNumb, oldInitNumb;
        private string[] keyCompID = {"", "", "Water"};
        public NeqSimNETService neqsimService;
        public int nNumComp = 2;
        public int numPhases = 3;

        public int[] oilFractionIDs;

        private readonly string[] phaseL = {"Vapor", "Liquid", "Liquid2"};

        public string[] properties;
        private string[] stateOfA = {"Vapor", "Liquid", "Liquid"};
        public string[] twoProp;

        public NeqSimNETClientCO10()
        {
        }

        public NeqSimNETClientCO10(string packageDesc)
        {
            var packages = packageDesc.Split(' ');
            var package = packages[0];

            neqsimService = new NeqSimNETService();
            neqsimService.readFluidFromGQIT(Convert.ToInt32(package));
            neqsimService.setPackageID(Convert.ToInt32(package));

            properties = neqsimService.readCapeOpenProperties10();

            nNumComp = neqsimService.getNumberOfComponents();
            twoProp = null; // new string[] { "surfaceTension" };//, "kvalue", "logKvalue" };
            oilFractionIDs = neqsimService.getOilFractionIDs();
            constPropList = new[]
            {
                "liquidDensityAt25C", "molecularWeight", "normalBoilingPoint", "acentricFactor", "criticalPressure",
                "criticalTemperature", "criticalVolume"
            };
            componentName = packageDesc;
            componentDescription = packageDesc; //"NeqSim thermo package";
        }


        public virtual string ComponentDescription
        {
            get => componentDescription;
            set => componentDescription = value;
        }

        public virtual string ComponentName
        {
            get => componentName;
            set => componentName = value;
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


        public void GetComponentList(ref object compIds, ref object formulae, ref object names, ref object boilTemps,
            ref object molwts, ref object casnos)
        {
            neqsimService.GetCompoundList(ref compIds, ref formulae, ref names, ref boilTemps, ref molwts, ref casnos);
        }

        public void CalcEquilibrium(object specification1, string specification2, object name)
        {
        }

        public void CalcProp(object material, object props, object phaseLabels, string name2)
        {
            try
            {
                var lables = (string[]) phaseLabels;
                for (var i = 0; i < lables.Length; i++)
                {
                    if (lables[i].Equals("Overall"))
                        CalcSinglePhaseProp((ICapeThermoMaterialObject) material, props, lables[i]);
                    CalcSinglePhaseProp((ICapeThermoMaterialObject) material, props, lables[i]);
                }
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }
        }

        public object GetComponentConstant(object material, object props)
        {
            var numberOfComponents = 0;

            numberOfComponents = GetNumCompounds();

            var propsName = (string[]) props;
            var tempString = (string[]) props;
            var propLength = tempString.Length;
            if (propLength == 0) return null;

            var propsVal = new object[numberOfComponents];

            for (var i = 0; i < propLength; i++)
            for (var j = 0; j < numberOfComponents; j++)
            {
                if (propsName[i].Equals("criticalVolume")) propsVal[j] = neqsimService.getCriticalVolumes()[j];
                if (propsName[i].Equals("criticalTemperature"))
                    propsVal[j] = neqsimService.getCriticalTemperatures()[j];
                if (propsName[i].Equals("criticalPressure")) propsVal[j] = neqsimService.getCriticalPressures()[j];
                if (propsName[i].Equals("acentricFactor")) propsVal[j] = neqsimService.getAcentricFactor()[j];
                if (propsName[i].Equals("normalBoilingPoint")) propsVal[j] = neqsimService.getNormalBoilingPoint()[j];
                if (propsName[i].Equals("molecularWeight")) propsVal[j] = neqsimService.getMolecularWeight()[j];
                if (propsName[i].Equals("liquidDensityAt25C")) propsVal[j] = neqsimService.getLiquidDensityAt25C()[j];
            }

            return propsVal;
        }


        public void CalcSinglePhaseProp(ICapeThermoMaterialObject material, object props, string phaseLabel)
        {
            try
            {
                var tempString = (string[]) props;
                var length = tempString.Length;
                double temperature = 0;
                double pressure = 0;
                object composition = null;
                try
                {
                    temperature =
                        material.GetProp("temperature", "Overall", null, "Mixture", "")[0]; //"ref composition);
                    pressure = material.GetProp("pressure", "Overall", null, "Mixture", "")[0];
                    composition = material.GetProp("fraction", phaseLabel, null, "Mixture", "mole");
                }
                catch (Exception e)
                {
                    var w = e.Message;
                    throw e;
                }

                if (pressure > 2e8)
                {
                    pressure = 2e8;
                    return;
                }


                var doInit = true;

                for (var i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("logFugacityCoefficient") || tempString[i].Equals("fugacityCoefficient") ||
                        tempString[i].StartsWith("molecularWeight") || tempString[i].Equals("volume.Dpressure") ||
                        tempString[i].Equals("density") || tempString[i].Equals("compressibilityFactor") ||
                        tempString[i].Equals("density.Dpressure") || tempString[i].Equals("density.Dmoles") ||
                        tempString[i].Equals("volume") || tempString[i].Equals("helmholtzEnergy") ||
                        tempString[i].Equals("gibbsEnergy") || tempString[i].Equals("viscosity") ||
                        tempString[i].Equals("thermalConductivity")) initNumb = 1;
                    else if (tempString[i].Equals("logFugacityCoefficient.Dmoles")) initNumb = 3;
                    else initNumb = 2;

                    if (initNumb > oldInitNumb || neqsimService.checkIfInitNeed(temperature, pressure / 1.0e5,
                            (double[]) composition, phaseLabel))
                    {
                        oldInitNumb = initNumb;
                        neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[]) composition);
                        neqsimService.init(phaseLabel, initNumb);
                        doInit = false;
                    }
                    else
                    {
                        doInit = false;
                    }
                }

                for (var i = 0; i < length; i++)
                {
                    if (tempString[i].Equals("fugacityCoefficient"))
                    {
                        var fugCoefs = neqsimService.getFugacityCoefficients(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefs);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient"))
                    {
                        //double[] lnPhiTemp = new double[nNumComp];
                        var lnPhiTemp = neqsimService.getLogFugacityCoefficients(phaseLabel, doInit);
                        //for (int k = 0; k < fugCoef.Length; k++)
                        //{
                        //    lnPhiTemp[k] = Math.Log(fugCoef[k]);
                        // }
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", lnPhiTemp);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dtemperature"))
                    {
                        var fugCoefsdT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, doInit);

                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefsdT);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dpressure"))
                    {
                        var fugCoefsdP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", fugCoefsdP);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dmoles"))
                    {
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole",
                            neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, doInit));
                        continue;
                    }

                    if (tempString[i].Equals("density"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("compressibilityFactor"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getCompressibilityFactor(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dtemperature"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensitydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dpressure"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensitydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight"))
                    {
                        var molWtrer = new double[1];
                        molWtrer[0] = neqsimService.getMolecularWeight(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", molWtrer);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight.Dtemperature") ||
                        tempString[i].Equals("molecularWeight.Dpressure"))
                    {
                        var molWtrer = new double[1];
                        molWtrer[0] = 0.0;
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", molWtrer);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight.Dmoles"))
                    {
                        var Mdn = neqsimService.getMolecularWeightdN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", Mdn);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getDensitydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                        continue;
                    }

                    if (tempString[i].Equals("volume"))
                    {
                        var volume = new double[1];
                        volume[0] = neqsimService.getVolume(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volume);
                        continue;
                    }

                    if (tempString[i].Equals("volume.Dtemperature"))
                    {
                        var volumedT = new double[1];
                        volumedT[0] = neqsimService.getVolumedT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volumedT);
                    }
                    else if (tempString[i].Equals("volume.Dpressure"))
                    {
                        var volumedP = new double[1];
                        volumedP[0] = neqsimService.getVolumedP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", volumedP);
                    }
                    else if (tempString[i].Equals("heatCapacity"))
                    {
                        var heatCapacityCp = new double[1];
                        heatCapacityCp[0] = neqsimService.getHeatCapacityCp(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", heatCapacityCp);
                    }
                    else if (tempString[i].Equals("heatCapacityCv"))
                    {
                        var heatCapacityCv = new double[1];
                        heatCapacityCv[0] = neqsimService.getHeatCapacityCv(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", heatCapacityCv);
                    }
                    else if (tempString[i].Equals("enthalpy"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                    }
                    else if (tempString[i].Equals("enthalpy.Dtemperature"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                    }
                    else if (tempString[i].Equals("enthalpy.Dpressure"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpy);
                    }
                    else if (tempString[i].Equals("enthalpy.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getEnthalpydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                    }
                    else if (tempString[i].Equals("entropy"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                    }
                    else if (tempString[i].Equals("entropy.Dpressure"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydP(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                    }
                    else if (tempString[i].Equals("entropy.Dtemperature"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", entropy);
                    }
                    else if (tempString[i].Equals("entropy.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getEntropydN(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", enthalpydn);
                    }
                    else if (tempString[i].Equals("helmholtzEnergy"))
                    {
                        var H = new double[1];
                        H[0] = neqsimService.getHelmholtzEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", H);
                    }
                    else if (tempString[i].Equals("jouleThomsonCoefficient"))
                    {
                        var JT = new double[1];
                        JT[0] = neqsimService.getJouleThomsonCoefficient(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", JT);
                    }
                    else if (tempString[i].Equals("gibbsEnergy"))
                    {
                        var gibbsEnergy = new double[1];
                        gibbsEnergy[0] = neqsimService.getGibbsEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", gibbsEnergy);
                    }
                    else if (tempString[i].Equals("energy"))
                    {
                        var internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergy(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", internalEnergy);
                    }
                    else if (tempString[i].Equals("energy.Dtemperature"))
                    {
                        var internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergydT(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", internalEnergy);
                    }
                    else if (tempString[i].Equals("speedOfSound"))
                    {
                        var speedOfSound = new double[1];
                        speedOfSound[0] = neqsimService.getSpeedOfSound(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", speedOfSound);
                    }

                    else if (tempString[i].Equals("viscosity"))
                    {
                        var viscosity = new double[1];
                        viscosity[0] = neqsimService.getViscosity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", viscosity);
                    }
                    else if (tempString[i].Equals("thermalConductivity"))
                    {
                        var conductivity = new double[1];
                        conductivity[0] = neqsimService.getThermalConductivity(phaseLabel, doInit);
                        material.SetProp(tempString[i], phaseLabel, null, "Mixture", "Mole", conductivity);
                    }
                }
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }
        }

        public virtual int GetNumCompounds()
        {
            return nNumComp;
        }

        public object GetComponentConstant2(object props, object compIds)
        {
            var numberOfComponents = 0;

            if (compIds == null) numberOfComponents = GetNumCompounds();
            else
                numberOfComponents = ((string[]) compIds).Length;

            var propsName = (string[]) props;
            var tempString = (string[]) props;
            var propLength = tempString.Length;
            if (propLength == 0) return null;

            var propsVal = new object[numberOfComponents];

            for (var i = 0; i < propLength; i++)
            for (var j = 0; j < numberOfComponents; j++)
            {
                if (propsName[i].Equals("criticalVolume")) propsVal[j] = neqsimService.getCriticalVolumes()[j];
                if (propsName[i].Equals("criticalTemperature"))
                    propsVal[j] = neqsimService.getCriticalTemperatures()[j];
                if (propsName[i].Equals("criticalPressure")) propsVal[j] = neqsimService.getCriticalPressures()[j];
                if (propsName[i].Equals("acentricFactor")) propsVal[j] = neqsimService.getAcentricFactor()[j];
                if (propsName[i].Equals("normalBoilingPoint")) propsVal[j] = neqsimService.getNormalBoilingPoint()[j];
                if (propsName[i].Equals("molecularWeight")) propsVal[j] = neqsimService.getMolecularWeight()[j];
                if (propsName[i].Equals("liquidDensityAt25C")) propsVal[j] = neqsimService.getLiquidDensityAt25C()[j];
            }

            return propsVal;
        }
    }
}