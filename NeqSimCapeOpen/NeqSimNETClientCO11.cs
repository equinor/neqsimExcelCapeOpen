using System;
using NeqSimNET;

namespace CapeOpenThermo
{
    public class NeqSimNETClientCO11 : NeqSimBasePackage
    {
        private int initNumb, oldInitNumb;

        public NeqSimNETClientCO11()
        {
        }


        public NeqSimNETClientCO11(string packageDesc)
        {
            var packages = packageDesc.Split(' ');
            var package = packages[0];

            neqsimService = new NeqSimNETService();
            neqsimService.readFluidFromGQIT(Convert.ToInt32(package));
            neqsimService.setPackageID(Convert.ToInt32(package));

            neqsimService.addCapeOpenProperty("viscosity");
         //   neqsimService.addCapeOpenProperty("thermalConductivity");
            properties = neqsimService.readCapeOpenProperties11();

            nNumComp = neqsimService.getNumberOfComponents();
            twoProp = null;//new[]
               // {"surfaceTension"}; // null; new string[] { "surfaceTension" };//, "kvalue", "logKvalue" }; =null;//
            oilFractionIDs = neqsimService.getOilFractionIDs();
           constPropList = new[]
            {
                "liquidDensityAt25C", "molecularWeight", "normalBoilingPoint", "acentricFactor", "criticalPressure",
                "criticalTemperature", "criticalVolume"
            };
            componentName = packageDesc;
            componentDescription = packageDesc;
        }

        public NeqSimNETClientCO11(string package, string packageName)
        {
            neqsimService = new NeqSimNETService();
            neqsimService.readFluidFromGQIT(package);
            // neqsimService.setPackageID(Convert.ToInt32(package));

            //neqsimService.addCapeOpenProperty("viscosity");
            //neqsimService.addCapeOpenProperty("thermalConductivity");
            properties = neqsimService.readCapeOpenProperties11();

            nNumComp = neqsimService.getNumberOfComponents();
            twoProp = null;//new[] {"surfaceTension"}; // null; new string[] { "surfaceTension" };//, "kvalue", "logKvalue" }; =null;//
            oilFractionIDs = neqsimService.getOilFractionIDs();
            componentName = packageName;
            componentDescription = packageName;
            constPropList = new[]
            {
                "liquidDensityAt25C", "molecularWeight", "normalBoilingPoint", "acentricFactor", "criticalPressure",
                "criticalTemperature", "criticalVolume"
            };
        }


        public override void GetCompoundList(ref object compIds, ref object formulae, ref object names,
            ref object boilTemps, ref object molwts, ref object casnos)
        {
            neqsimService.GetCompoundList(ref compIds, ref formulae, ref names, ref boilTemps, ref molwts, ref casnos);
        }

        public override int GetNumCompounds()
        {
            return neqsimService.getNumberOfComponents();
        }


        public override object GetCompoundConstant(object props, object compIds)
        {
            var numberOfComponents = 0;

            if (compIds == null)
            {
                numberOfComponents = GetNumCompounds();
                compIds = neqsimService.getComponentIDs();
            }
            else
            {
                numberOfComponents = ((string[]) compIds).Length;
            }

            var propsName = (string[]) props;
            var propLength = propsName.Length;
            if (propLength == 0) return null;

            var propsVal = new object[numberOfComponents];

            for (var j = 0; j < numberOfComponents; j++)
            {
                var stringID = ((string[]) compIds)[j];
                var value = new double[propLength];

                for (var i = 0; i < propLength; i++)
                    if (propsName[i].Equals("criticalVolume"))
                        value[i] = neqsimService.getCriticalVolume(stringID);
                    else if (propsName[i].Equals("criticalTemperature"))
                        value[i] = neqsimService.getCriticalTemperature(stringID);
                    else if (propsName[i].Equals("criticalPressure"))
                        value[i] = neqsimService.getCriticalPressure(stringID);
                    else if (propsName[i].Equals("acentricFactor"))
                        value[i] = neqsimService.getAcentricFactor(stringID);
                    else if (propsName[i].Equals("normalBoilingPoint"))
                        value[i] = neqsimService.getNormalBoilingPoint(stringID);
                    else if (propsName[i].Equals("molecularWeight"))
                        value[i] = neqsimService.getMolecularWeight(stringID);
                    else if (propsName[i].Equals("liquidDensityAt25C"))
                        value[i] = neqsimService.getLiquidDensityAt25C(stringID);
                if (propLength == 1) propsVal[j] = value[0];
                else propsVal[j] = value;
            }

            return propsVal;
        }

        public override void CalcAndGetLnPhi(string phaseLabel, double temperature, double pressure, object moleNumbers,
            int fFlags, ref object lnPhi, ref object lnPhiDT, ref object lnPhiDP, ref object lnPhiDn)
        {
            switch (fFlags)
            {
                case 0:
                    return;
                case 1:
                    initNumb = 1;
                    break;
                case 2:
                    initNumb = 2;
                    break;
                case 4:
                    initNumb = 2;
                    break;
                case 8:
                    initNumb = 3;
                    break;
            }

            if (initNumb > oldInitNumb ||
                neqsimService.checkIfInitNeed(temperature, pressure / 1.0e5, (double[]) moleNumbers, phaseLabel))
            {
                oldInitNumb = initNumb;
                neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[]) moleNumbers);
                neqsimService.init(phaseLabel, initNumb);
            }

            /*  if (!neqsimService.PhaseExist)
              {
                  throw new PhaseDoesNotExcistExeption("phase noes not exsist");
              }
              */
            switch (fFlags)
            {
                case 1:
                    lnPhi = neqsimService.getLogFugacityCoefficients(phaseLabel, false);
                    break;
                case 2:
                    lnPhiDT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, false);
                    break;
                case 4:
                    lnPhiDP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, false);
                    break;
                case 8:
                    lnPhiDn = neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, false);
                    break;
            }
        }

        public override void CalcSinglePhaseProp(object props, string phaseLabel)
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
                    material.GetTPFraction(phaseLabel, ref temperature, ref pressure, ref composition);
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
                    if (tempString[i].Equals("logFugacityCoefficient") || tempString[i].Equals("volume") ||
                        tempString[i].Equals("fugacityCoefficient") || tempString[i].StartsWith("molecularWeight") ||
                        tempString[i].Equals("volume.Dpressure") || tempString[i].Equals("density") ||
                        tempString[i].Equals("compressibilityFactor") || tempString[i].Equals("density.Dpressure") ||
                        tempString[i].Equals("density.Dmoles") || tempString[i].Equals("helmholtzEnergy") ||
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
                        /*
                        if (!neqsimService.PhaseExist)
                        {
                            throw new PhaseDoesNotExcistExeption("phase noes not exsist");
                        }
                        */
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
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefs);
                        continue;
                    }

                    if (tempString[i].Equals("volume"))
                    {
                        var volume = new double[1];
                        volume[0] = neqsimService.getVolume(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volume);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient"))
                    {
                        var lnPhiTemp = neqsimService.getLogFugacityCoefficients(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiTemp);
                        continue;
                    }

                    if (tempString[i].Equals("enthalpy"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                        continue;
                    }

                    if (tempString[i].Equals("entropy"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dtemperature"))
                    {
                        var fugCoefsdT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefsdT);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dpressure"))
                    {
                        var fugCoefsdP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefsdP);
                        continue;
                    }

                    if (tempString[i].Equals("logFugacityCoefficient.Dmoles"))
                    {
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "",
                            neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, doInit));
                        continue;
                    }

                    if (tempString[i].Equals("density"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("compressibilityFactor"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getCompressibilityFactor(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dtemperature"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensitydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dpressure"))
                    {
                        var density = new double[1];
                        density[0] = neqsimService.getDensitydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight"))
                    {
                        var molWtrer = new double[1];
                        molWtrer[0] = neqsimService.getMolecularWeight(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", molWtrer);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight.Dtemperature") ||
                        tempString[i].Equals("molecularWeight.Dpressure"))
                    {
                        var molWtrer = new double[1];
                        molWtrer[0] = 0.0;
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", molWtrer);
                        continue;
                    }

                    if (tempString[i].Equals("molecularWeight.Dmoles"))
                    {
                        var Mdn = neqsimService.getMolecularWeightdN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", Mdn);
                        continue;
                    }

                    if (tempString[i].Equals("density.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getDensitydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                        continue;
                    }

                    if (tempString[i].Equals("volume.Dtemperature"))
                    {
                        var volumedT = new double[1];
                        volumedT[0] = neqsimService.getVolumedT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volumedT);
                    }
                    else if (tempString[i].Equals("volume.Dpressure"))
                    {
                        var volumedP = new double[1];
                        volumedP[0] = neqsimService.getVolumedP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volumedP);
                    }
                    else if (tempString[i].Equals("heatCapacityCp"))
                    {
                        var heatCapacityCp = new double[1];
                        heatCapacityCp[0] = neqsimService.getHeatCapacityCp(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", heatCapacityCp);
                    }
                    else if (tempString[i].Equals("heatCapacityCv"))
                    {
                        var heatCapacityCv = new double[1];
                        heatCapacityCv[0] = neqsimService.getHeatCapacityCv(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", heatCapacityCv);
                    }

                    else if (tempString[i].Equals("enthalpy.Dtemperature"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    }
                    else if (tempString[i].Equals("enthalpy.Dpressure"))
                    {
                        var enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                    }
                    else if (tempString[i].Equals("enthalpy.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getEnthalpydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                    }
                    else if (tempString[i].Equals("entropy.Dpressure"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                    }
                    else if (tempString[i].Equals("entropy.Dtemperature"))
                    {
                        var entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                    }
                    else if (tempString[i].Equals("entropy.Dmoles"))
                    {
                        var enthalpydn = neqsimService.getEntropydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                    }
                    else if (tempString[i].Equals("helmholtzEnergy"))
                    {
                        var helmlE = new double[1];
                        helmlE[0] = neqsimService.getHelmholtzEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", helmlE);
                    }
                    else if (tempString[i].Equals("jouleThomsonCoefficient"))
                    {
                        var JT = new double[1];
                        JT[0] = neqsimService.getJouleThomsonCoefficient(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", JT);
                    }
                    else if (tempString[i].Equals("gibbsEnergy"))
                    {
                        var gibbsEnergy = new double[1];
                        gibbsEnergy[0] = neqsimService.getGibbsEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", gibbsEnergy);
                    }
                    else if (tempString[i].Equals("internalEnergy"))
                    {
                        var internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", internalEnergy);
                    }
                    else if (tempString[i].Equals("internalEnergy.Dtemperature"))
                    {
                        var internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", internalEnergy);
                    }
                    else if (tempString[i].Equals("speedOfSound"))
                    {
                        var speedOfSound = new double[1];
                        speedOfSound[0] = neqsimService.getSpeedOfSound(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", speedOfSound);
                    }

                    else if (tempString[i].Equals("viscosity"))
                    {
                        var viscosity = new double[1];
                        viscosity[0] = neqsimService.getViscosity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", viscosity);
                    }
                    else if (tempString[i].Equals("thermalConductivity"))
                    {
                        var conductivity = new double[1];
                        conductivity[0] = neqsimService.getThermalConductivity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", conductivity);
                    }
                }
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }
        }

        public override void CalcEquilibrium(object specification1, object specification2, string name)
        {
            var spec1 = (string[]) specification1;
            var spec2 = (string[]) specification2;
            if (spec1[1] == null) spec1[1] = "";
            if (spec2[1] == null) spec2[1] = "";


            spec1[0] = spec1[0].ToLower();
            spec2[0] = spec2[0].ToLower();
            spec1[1] = spec1[1].ToLower();
            spec2[1] = spec2[1].ToLower();
            spec1[2] = spec1[2].ToLower();
            spec2[2] = spec2[2].ToLower();
            try
            {
                neqsimService.initFlashCalc();

                if (spec1[0].Equals("temperature") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                    if (spec2[0].Equals("pressure") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                        TPflash();

                if (spec1[0].Equals("enthalpy") || spec1[0].Equals("entropy"))
                {
                    //  CalcEquilibrium(spec2, spec1, name);
                }


                if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                    if (spec2[0].Equals("enthalpy") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                        PHflash();

                if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                    if (spec2[0].Equals("entropy") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                        PSflash();

                if (spec1[0].Equals("temperature") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                    if (spec2[0].Equals("phasefraction") && spec2[1].Equals("mole") && spec2[2].Equals("vapor"))
                        phaseFractionFlash("vapor");
            }
            finally
            {
                neqsimService.endFlashCalc();
            }
        }

        public void phaseFractionFlash(string phase)
        {
            object temperature = 298.15;
            object composition = null;

            object phaseFraction = 0;
            object pressure = 10.0;


            //material.SetOverallProp("temperature", "", valTemp);
            try
            {
                material.GetSinglePhaseProp("phaseFraction", phase, "Mole", ref phaseFraction);
                material.GetOverallProp("temperature", "", ref temperature);
                material.GetOverallProp("fraction", "Mole", ref composition);

                //  material.GetSinglePhaseProp("temperature", phase, "", ref temperature);
                //  material.GetSinglePhaseProp("fraction", phase, "Mole", ref composition);
                // material.GetOverallProp("phaseFraction", "Mole", ref phaseFraction);
                // material.GetOverallProp("temperature", "", ref temperature);
                // material.GetOverallProp("fraction", "Mole", ref composition);
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }

            neqsimService.setTPFractionFlash(((double[]) temperature)[0], (double) pressure, (double[]) composition);

            neqsimService.phaseFractionFlash(phase, ((double[]) phaseFraction)[0]);
            pressure = neqsimService.getPressure() * 1e5;

            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();
            var presentPhasesStatus = new int[((string[]) presentPhases).Length];
            for (var i = 0; i < ((string[]) presentPhases).Length; i++) presentPhasesStatus[i] = 1;
            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            for (var i = 0; i < ((string[]) presentPhases).Length; i++)
            {
                var phaseCompositions = neqsimService.getPhaseComposition(i);
                var val = new double[1];
                val[0] = (double) temperature;
                material.SetSinglePhaseProp("temperature", ((string[]) presentPhases)[i], "", val);
                val[0] = (double) pressure;
                material.SetSinglePhaseProp("pressure", ((string[]) presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[]) presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[]) presentPhases)[i], "Mole", val);
            }
        }

        public void TPflash()
        {
            double temperature = 0;
            double pressure = 0;
            object composition = null;
            try
            {
                material.GetOverallTPFraction(ref temperature, ref pressure, ref composition);
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }

            neqsimService.setTPFractionFlash(temperature, pressure / 1.0e5, (double[]) composition);
            neqsimService.TPflash();
            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();
            var presentPhasesStatus = new int[((string[]) presentPhases).Length];
            for (var i = 0; i < ((string[]) presentPhases).Length; i++) presentPhasesStatus[i] = 1;
            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            for (var i = 0; i < ((string[]) presentPhases).Length; i++)
            {
                var phaseCompositions = neqsimService.getPhaseComposition(i);
                var val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[]) presentPhases)[i], "", val);
                val[0] = pressure;
                material.SetSinglePhaseProp("pressure", ((string[]) presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[]) presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[]) presentPhases)[i], "Mole", val);
            }
        }

        public void PHflash()
        {
            var temperature = 298.15;
            object composition = null;

            object enthalpy = 0;
            object pressure = new double[1];


            //material.SetOverallProp("temperature", "", valTemp);
            try
            {
                material.GetOverallProp("enthalpy", "Mole", ref enthalpy);
                material.GetOverallProp("pressure", "", ref pressure);
                material.GetOverallProp("fraction", "Mole", ref composition);
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }

            var totEnthalpy = ((double[]) enthalpy)[0];
            neqsimService.setTPFractionFlash(temperature, ((double[]) pressure)[0] / 1.0e5, (double[]) composition);
            neqsimService.PHflash(totEnthalpy);
            temperature = neqsimService.getTemperature();

            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();

            var presentPhasesStatus = new int[((string[]) presentPhases).Length];
            for (var i = 0; i < ((string[]) presentPhases).Length; i++) presentPhasesStatus[i] = 1;

            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            var val = new double[1];
            //
            val[0] = temperature;
            material.SetOverallProp("temperature", "", val);


            for (var i = 0; i < ((string[]) presentPhases).Length; i++)
            {
                var phaseCompositions = neqsimService.getPhaseComposition(i);
                //material.SetPresentPhases(phases, phases);
                val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[]) presentPhases)[i], "", val);
                val[0] = ((double[]) pressure)[0];
                material.SetSinglePhaseProp("pressure", ((string[]) presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[]) presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[]) presentPhases)[i], "Mole", val);
            }
        }

        public void PSflash()
        {
            var temperature = 298.15;
            object composition = null;

            object entropy = 0;
            object pressure = new double[1];


            try
            {
                material.GetOverallProp("entropy", "Mole", ref entropy);
                material.GetOverallProp("pressure", "", ref pressure);
                material.GetOverallProp("fraction", "Mole", ref composition);
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }

            //material.SetOverallProp("enthalpy", "Mole", ref enthalpy);
            var totEntropy = ((double[]) entropy)[0];
            neqsimService.setTPFractionFlash(temperature, ((double[]) pressure)[0] / 1.0e5, (double[]) composition);
            neqsimService.PSflash(totEntropy);
            temperature = neqsimService.getTemperature();
            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
            // material.ClearAllProps();

            var presentPhasesStatus = new int[((string[]) presentPhases).Length];
            for (var i = 0; i < ((string[]) presentPhases).Length; i++) presentPhasesStatus[i] = 1;

            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            var val = new double[1];
            //
            val[0] = temperature;
            material.SetOverallProp("temperature", "", val);

            for (var i = 0; i < ((string[]) presentPhases).Length; i++)
            {
                var phaseCompositions = neqsimService.getPhaseComposition(i);
                //material.SetPresentPhases(phases, phases);
                val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[]) presentPhases)[i], "", val);
                val[0] = ((double[]) pressure)[0];
                material.SetSinglePhaseProp("pressure", ((string[]) presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[]) presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[]) presentPhases)[i], "Mole", val);
            }
        }

        public override bool CheckEquilibriumSpec(object specification1, object specification2, string name)
        {
            var spec1 = (string[]) specification1;
            var spec2 = (string[]) specification2;

            if (spec1[1] == null) spec1[1] = "";
            if (spec2[1] == null) spec2[1] = "";


            spec1[0] = spec1[0].ToLower();
            spec2[0] = spec2[0].ToLower();

            if (spec1[0].Equals("temperature") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
                if (spec2[0].Equals("pressure") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                    return true; //support TPflash
            if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
                if (spec2[0].Equals("enthalpy") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                    return true; //support PHflash
            if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
                if (spec2[0].Equals("entropy") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                    return true; //support PSflash
            return false;
        }

        public override void CalcTwoPhaseProp(object props, object phaseLabels)
        {
            var propertyNames = (string[]) props;
            var phaseNames = (string[]) phaseLabels;
            var numberOfTwoPhaseProperties = propertyNames.Length;
            double temperature = 273.15, pressure = 1.0e5;
            object composition1 = new double[neqsimService.getNumberOfComponents()];
            object composition2 = new double[neqsimService.getNumberOfComponents()];


            try
            {
                material.GetTPFraction(phaseNames[0].ToLower(), ref temperature, ref pressure, ref composition1);
                material.GetTPFraction(phaseNames[1].ToLower(), ref temperature, ref pressure, ref composition2);
            }
            catch (Exception e)
            {
                var errorMessage = e.ToString();
                throw e;
            }

            neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[]) composition1,
                (double[]) composition2);

            for (var i = 0; i < numberOfTwoPhaseProperties; i++)
                if (propertyNames[i].Equals("surfaceTension"))
                {
                    var surfTens = new double[1];
                    surfTens[0] = neqsimService.getSurfaceTension(phaseNames[0], phaseNames[1]) * 1e3;
                    material.SetTwoPhaseProp(propertyNames[i], phaseLabels, "", surfTens);
                }
        }
    }
}