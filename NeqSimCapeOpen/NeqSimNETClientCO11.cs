using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NeqSimNET;
using CAPEOPEN110;

namespace CapeOpenThermo
{
    public class NeqSimNETClientCO11 : NeqSimBasePackage
    {

        public NeqSimNETService neqsimService = null;
        int initNumb = 0, oldInitNumb = 0;
       


        public NeqSimNETClientCO11()
        {
        }


        public NeqSimNETClientCO11(String package)
        {
            neqsimService = new NeqSimNETService();
            neqsimService.readFluidFromGQIT(Convert.ToInt32(package));
            neqsimService.setPackageID(Convert.ToInt32(package));

            neqsimService.addCapeOpenProperty("viscosity");
            neqsimService.addCapeOpenProperty("thermalConductivity");
            properties = neqsimService.readCapeOpenProperties11();

            nNumComp = neqsimService.getNumberOfComponents();
            twoProp = new string[] { "surfaceTension" };// null; new string[] { "surfaceTension" };//, "kvalue", "logKvalue" }; =null;//
            oilFractionIDs = neqsimService.getOilFractionIDs();
            ComponentName = "NeqSim thermo package";
            ComponentDescription = "NeqSim thermo package";
            constPropList = new String[] { "liquidDensityAt25C", "molecularWeight", "normalBoilingPoint", "acentricFactor", "criticalPressure", "criticalTemperature", "criticalVolume" };

        }

        public override void GetCompoundList(ref object compIds, ref object formulae, ref object names, ref object boilTemps, ref object molwts, ref object casnos)
        {
            neqsimService.GetCompoundList(ref compIds, ref formulae, ref names, ref boilTemps, ref molwts, ref casnos);
        }

        public override int GetNumCompounds()
        {
            return neqsimService.getNumberOfComponents();
        }


        public override object GetCompoundConstant(object props, object compIds)
        {
            int numberOfComponents = 0;

            if (compIds == null)
            {
                numberOfComponents = GetNumCompounds();
                compIds = neqsimService.getComponentIDs();
            }
            else
            {
                numberOfComponents = ((string[])compIds).Length;
            }

            string[] propsName = (string[])props;
            int propLength = propsName.Length;
            if (propLength == 0) return null;

            object[] propsVal = new object[numberOfComponents];
            
                for (int j = 0; j < numberOfComponents; j++)
                {
                string stringID = ((string[])compIds)[j];
                double[] value = new double[propLength];

                for (int i = 0; i < propLength; i++)
                {
                        if (propsName[i].Equals("criticalVolume")) value[i] = neqsimService.getCriticalVolume(stringID);
                        else if (propsName[i].Equals("criticalTemperature")) value[i] = neqsimService.getCriticalTemperature(stringID);
                        else if (propsName[i].Equals("criticalPressure")) value[i] = neqsimService.getCriticalPressure(stringID);
                        else if (propsName[i].Equals("acentricFactor")) value[i] = neqsimService.getAcentricFactor(stringID);
                        else if (propsName[i].Equals("normalBoilingPoint")) value[i] = neqsimService.getNormalBoilingPoint(stringID);
                        else if (propsName[i].Equals("molecularWeight")) value[i] = neqsimService.getMolecularWeight(stringID);
                        else if (propsName[i].Equals("liquidDensityAt25C")) value[i] = neqsimService.getLiquidDensityAt25C(stringID);
                    }
                    if (propLength == 1) propsVal[j] = value[0];
                    else propsVal[j] = value;
                }
            return propsVal;
        }

        public override void CalcAndGetLnPhi(string phaseLabel, double temperature, double pressure, object moleNumbers, int fFlags, ref object lnPhi, ref object lnPhiDT, ref object lnPhiDP, ref object lnPhiDn)
        {
            if (fFlags == 0) return;
            
            neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[])moleNumbers);

            if (fFlags >= 8)
            {
                neqsimService.init(phaseLabel, 3);
                oldInitNumb = 3;
            }
            else if (fFlags >= 2)
            {
                neqsimService.init(phaseLabel, 2);
                oldInitNumb = 2;
            }
            else
            {
                neqsimService.init(phaseLabel, 1);
                oldInitNumb = 3;
            }

            if (fFlags >= 8) lnPhiDn = neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, false);
            if (fFlags >= 4) lnPhiDP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, false);
            if (fFlags >= 2) lnPhiDT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, false);
            if (fFlags >= 1) lnPhi = neqsimService.getLogFugacityCoefficients(phaseLabel, false);
            return;
        }

        public override void CalcSinglePhaseProp(object props, string phaseLabel)
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
                    material.GetTPFraction(phaseLabel, ref temperature, ref pressure, ref composition);
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
                    if (tempString[i].Equals("logFugacityCoefficient") || tempString[i].Equals("volume") || tempString[i].Equals("fugacityCoefficient") || tempString[i].StartsWith("molecularWeight") || tempString[i].Equals("volume.Dpressure") || tempString[i].Equals("density") || tempString[i].Equals("compressibilityFactor") || tempString[i].Equals("density.Dpressure") || tempString[i].Equals("density.Dmoles") || tempString[i].Equals("helmholtzEnergy") || tempString[i].Equals("gibbsEnergy") || tempString[i].Equals("viscosity") || tempString[i].Equals("thermalConductivity")) initNumb = 1;
                    else if (tempString[i].Equals("logFugacityCoefficient.Dmoles")) initNumb = 3;
                    else initNumb = 2;

                    if(initNumb > oldInitNumb || (neqsimService.checkIfInitNeed(temperature, pressure / 1.0e5, (double[])composition, phaseLabel)))
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
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefs);
                        continue;
                    }
                    else if (tempString[i].Equals("volume"))
                    {
                        double[] volume = new double[1];
                        volume[0] = neqsimService.getVolume(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volume);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient"))
                    {
                        double[] lnPhiTemp = neqsimService.getLogFugacityCoefficients(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", lnPhiTemp);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                        continue;
                    }

                    else if (tempString[i].Equals("entropy"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dtemperature"))
                    {
                        double[] fugCoefsdT = neqsimService.getlogFugacityCoefficientsDtemperature(phaseLabel, doInit);

                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefsdT);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dpressure"))
                    {
                        double[] fugCoefsdP = neqsimService.getlogFugacityCoefficientsDpressure(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", fugCoefsdP);
                        continue;
                    }
                    else if (tempString[i].Equals("logFugacityCoefficient.Dmoles"))
                    {
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", neqsimService.getlogFugacityCoefficientsDmoles(phaseLabel, doInit));
                        continue;
                    }
                    else if (tempString[i].Equals("density"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("compressibilityFactor"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getCompressibilityFactor(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dtemperature"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensitydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dpressure"))
                    {
                        double[] density = new double[1];
                        density[0] = neqsimService.getDensitydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", density);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight"))
                    {
                        double[] molWtrer = new double[1];
                        molWtrer[0] = neqsimService.getMolecularWeight(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", molWtrer);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight.Dtemperature") || tempString[i].Equals("molecularWeight.Dpressure"))
                    {
                        double[] molWtrer = new double[1];
                        molWtrer[0] = 0.0;
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", molWtrer);
                        continue;
                    }
                    else if (tempString[i].Equals("molecularWeight.Dmoles"))
                    {
                        double[] Mdn = neqsimService.getMolecularWeightdN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", Mdn);
                        continue;
                    }
                    else if (tempString[i].Equals("density.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getDensitydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                        continue;
                    }
                    if (tempString[i].Equals("volume.Dtemperature"))
                    {
                        double[] volumedT = new double[1];
                        volumedT[0] = neqsimService.getVolumedT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volumedT);
                        continue;
                    }
                    else if (tempString[i].Equals("volume.Dpressure"))
                    {
                        double[] volumedP = new double[1];
                        volumedP[0] = neqsimService.getVolumedP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", volumedP);
                        continue;
                     }
                    else if (tempString[i].Equals("heatCapacityCp"))
                    {
                        double[] heatCapacityCp = new double[1];
                        heatCapacityCp[0] = neqsimService.getHeatCapacityCp(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", heatCapacityCp);
                        continue;
                    }
                    else if (tempString[i].Equals("heatCapacityCv"))
                    {
                        double[] heatCapacityCv = new double[1];
                        heatCapacityCv[0] = neqsimService.getHeatCapacityCv(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", heatCapacityCv);
                        continue;
                    }
                   
                    else if (tempString[i].Equals("enthalpy.Dtemperature"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy.Dpressure"))
                    {
                        double[] enthalpy = new double[1];
                        enthalpy[0] = neqsimService.getEnthalpydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpy);
                        continue;
                    }
                    else if (tempString[i].Equals("enthalpy.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getEnthalpydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dpressure"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydP(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dtemperature"))
                    {
                        double[] entropy = new double[1];
                        entropy[0] = neqsimService.getEntropydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", entropy);
                        continue;
                    }
                    else if (tempString[i].Equals("entropy.Dmoles"))
                    {
                        double[] enthalpydn = neqsimService.getEntropydN(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", enthalpydn);
                        continue;
                    }
                    else if (tempString[i].Equals("helmholtzEnergy"))
                    {
                        double[] H = new double[1];
                        H[0] = neqsimService.getHelmholtzEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", H);
                        continue;
                    }
                    else if (tempString[i].Equals("jouleThomsonCoefficient"))
                    {
                        double[] JT = new double[1];
                        JT[0] = neqsimService.getJouleThomsonCoefficient(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", JT);
                        continue;
                    }
                    else if (tempString[i].Equals("gibbsEnergy"))
                    {
                        double[] gibbsEnergy = new double[1];
                        gibbsEnergy[0] = neqsimService.getGibbsEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", gibbsEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("internalEnergy"))
                    {
                        double[] internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergy(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", internalEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("internalEnergy.Dtemperature"))
                    {
                        double[] internalEnergy = new double[1];
                        internalEnergy[0] = neqsimService.getInternalEnergydT(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "Mole", internalEnergy);
                        continue;
                    }
                    else if (tempString[i].Equals("speedOfSound"))
                    {
                        double[] speedOfSound = new double[1];
                        speedOfSound[0] = neqsimService.getSpeedOfSound(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", speedOfSound);
                        continue;
                    }
                    
                    else if (tempString[i].Equals("viscosity"))
                    {
                        double[] viscosity = new double[1];
                        viscosity[0] = neqsimService.getViscosity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", viscosity);
                        continue;
                    }
                    else if (tempString[i].Equals("thermalConductivity"))
                    {
                        double[] conductivity = new double[1];
                        conductivity[0] = neqsimService.getThermalConductivity(phaseLabel, doInit);
                        material.SetSinglePhaseProp(tempString[i], phaseLabel, "", conductivity);
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

        public override void CalcEquilibrium(object specification1, object specification2, string name)
        {
            string[] spec1 = (string[])specification1;
            string[] spec2 = (string[])specification2;
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
                {
                    if (spec2[0].Equals("pressure") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                    {
                        TPflash();
                    }
                }

                if (spec1[0].Equals("enthalpy") || spec1[0].Equals("entropy"))
                {
                    CalcEquilibrium(spec2, spec1, name);
                }
               

                if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                {
                    if (spec2[0].Equals("enthalpy") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                    {
                        PHflash();
                    }
                }

                if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                {
                    if (spec2[0].Equals("entropy") && spec2[1].Equals("") && spec2[2].Equals("overall"))
                    {
                        PSflash();
                    }
                }

                if (spec1[0].Equals("temperature") && spec1[1].Equals("") && spec1[2].Equals("overall"))
                {
                    if (spec2[0].Equals("phasefraction") && spec2[1].Equals("mole") && spec2[2].Equals("vapor"))
                    {
                        phaseFractionFlash("vapor");
                    }
                }
            }
            finally
            {
                neqsimService.endFlashCalc();
            }

        }

        public void phaseFractionFlash(String phase)
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
                String w = e.Message;
                throw e;
            }
            
            neqsimService.setTPFractionFlash(((double[])temperature)[0], (double)pressure, (double[])composition);
            
            neqsimService.phaseFractionFlash(phase, ((double[])phaseFraction)[0]);
            pressure = neqsimService.getPressure()*1e5;

            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();
            int[] presentPhasesStatus = new int[((string[])presentPhases).Length];
            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                presentPhasesStatus[i] = 1;
            }
            material.SetPresentPhases(presentPhases, presentPhasesStatus);



            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                object phaseCompositions = neqsimService.getPhaseComposition(i);
                double[] val = new double[1];
                val[0] = (double)temperature;
                material.SetSinglePhaseProp("temperature", ((string[])presentPhases)[i], "", val);
                val[0] = (double)pressure;
                material.SetSinglePhaseProp("pressure", ((string[])presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[])presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[])presentPhases)[i], "Mole", val);
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
                String w = e.Message;
                throw e;
            }

            neqsimService.setTPFractionFlash(temperature, pressure / 1.0e5, (double[])composition);
            neqsimService.TPflash();
            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();
            int[] presentPhasesStatus = new int[((string[])presentPhases).Length];
            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                presentPhasesStatus[i] = 1;
            }
            material.SetPresentPhases(presentPhases, presentPhasesStatus);



            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                object phaseCompositions = neqsimService.getPhaseComposition(i);
                double[] val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[])presentPhases)[i], "", val);
                val[0] = pressure;
                material.SetSinglePhaseProp("pressure", ((string[])presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[])presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[])presentPhases)[i], "Mole", val);
            }

        }

        public void PHflash()
        {
            double temperature = 298.15;
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
                String w = e.Message;
                throw e;
            }

            double totEnthalpy = ((double[])enthalpy)[0];
            neqsimService.setTPFractionFlash(temperature, ((double[])pressure)[0] / 1.0e5, (double[])composition);
            neqsimService.PHflash(totEnthalpy);
            temperature = neqsimService.getTemperature();

            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
            //material.ClearAllProps();

            int[] presentPhasesStatus = new int[((string[])presentPhases).Length];
            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                presentPhasesStatus[i] = 1;
            }

            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            double[] val = new double[1];
            //
            val[0] = temperature;
            material.SetOverallProp("temperature", "", val);
            
            

            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                object phaseCompositions = neqsimService.getPhaseComposition(i);
                //material.SetPresentPhases(phases, phases);
                val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[])presentPhases)[i], "", val);
                val[0] = ((double[])pressure)[0];
                material.SetSinglePhaseProp("pressure", ((string[])presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[])presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[])presentPhases)[i], "Mole", val);
            }
        }

        public void PSflash()
        {
            double temperature = 298.15;
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
                String w = e.Message;
                throw e;
            }
            //material.SetOverallProp("enthalpy", "Mole", ref enthalpy);
            double totEntropy =  ((double[])entropy)[0];
            neqsimService.setTPFractionFlash(temperature, ((double[]) pressure)[0] / 1.0e5, (double[])composition);
            neqsimService.PSflash(totEntropy);
            temperature = neqsimService.getTemperature();
            //enum presentPhasesStatus { Cape_AtEquilibrium };
            object presentPhases = neqsimService.getPresentPhases();
           // material.ClearAllProps();

            int[] presentPhasesStatus = new int[((string[])presentPhases).Length];
            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                presentPhasesStatus[i] = 1;
            }

            material.SetPresentPhases(presentPhases, presentPhasesStatus);


            double[] val = new double[1];
            //
            val[0] = temperature;
            material.SetOverallProp("temperature", "", val);

            for (int i = 0; i < ((string[])presentPhases).Length; i++)
            {
                object phaseCompositions = neqsimService.getPhaseComposition(i);
                //material.SetPresentPhases(phases, phases);
                val = new double[1];
                val[0] = temperature;
                material.SetSinglePhaseProp("temperature", ((string[])presentPhases)[i], "", val);
                val[0] = ((double[])pressure)[0];
                material.SetSinglePhaseProp("pressure", ((string[])presentPhases)[i], "", val);
                material.SetSinglePhaseProp("fraction", ((string[])presentPhases)[i], "Mole", phaseCompositions);
                val[0] = neqsimService.getPhaseFraction(i);
                material.SetSinglePhaseProp("phaseFraction", ((string[])presentPhases)[i], "Mole", val);
            }
        }

        public override Boolean CheckEquilibriumSpec(object specification1, object specification2, string name)
        {
            string[] spec1 = (string[])specification1;
            string[] spec2 = (string[])specification2;

            if (spec1[1] == null) spec1[1] = "";
            if (spec2[1] == null) spec2[1] = "";



            spec1[0] = spec1[0].ToLower();
            spec2[0] = spec2[0].ToLower();

            if (spec1[0].Equals("temperature") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
            {
                if (spec2[0].Equals("pressure") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                {
                    return true; //support TPflash
                }
            }
            if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
            {
                if (spec2[0].Equals("enthalpy") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                {
                    return true; //support PHflash
                }
            }
            if (spec1[0].Equals("pressure") && spec1[1].Equals("") && spec1[2].Equals("Overall"))
            {
                if (spec2[0].Equals("entropy") && spec2[1].Equals("") && spec2[2].Equals("Overall"))
                {
                    return true; //support PSflash
                }
            }
            return false;
        }

        public override void CalcTwoPhaseProp(object props, object phaseLabels)
        {
            string[] propertyNames = (string[])props;
            string[] phaseNames = (string[])phaseLabels;
            int numberOfTwoPhaseProperties = propertyNames.Length;
            double temperature=273.15, pressure=1.0e5;
            object composition1 = new double[neqsimService.getNumberOfComponents()];
            object composition2 = new double[neqsimService.getNumberOfComponents()];


            try {
                material.GetTPFraction(phaseNames[0].ToLower(), ref temperature, ref pressure, ref composition1);
                material.GetTPFraction(phaseNames[1].ToLower(), ref temperature, ref pressure, ref composition2);
            }
            catch (Exception e)
            {
                String errorMessage = e.ToString();
                throw e;
            }

            neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[]) composition1, (double[])composition2);
            
            for (int i = 0; i < numberOfTwoPhaseProperties; i++)
            {
                if (propertyNames[i].Equals("surfaceTension"))
                    {
                        double[] surfTens = new double[1]; 
                        surfTens[0] = neqsimService.getSurfaceTension(phaseNames[0], phaseNames[1])*1e3;
                        material.SetTwoPhaseProp(propertyNames[i], phaseLabels, "", surfTens);
                    }
             }

            return;
        }
    }
}
