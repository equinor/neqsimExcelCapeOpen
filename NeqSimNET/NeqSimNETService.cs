using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using thermo.system;
using thermodynamicOperations;

namespace NeqSimNET
{
    public class NeqSimNETService
    {
        SystemInterface thermoSystem = (SystemInterface)new SystemSrkEos(298, 10);

        double[] oldMoleFraction = null;
        double oldTemperature = 1.00112;
        double oldPressure = 1.00112;
        int oldPhaseType = 2;
        public static int packageID = 0;
        


        public void test()
        {

            thermoSystem = thermoSystem.readObject(672);

            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();
            thermoSystem.display();

            thermoSystem.setTemperature(330);
            ops.TPflash();
            thermoSystem.display();
        }

        public void setPackageID(int id)
        {
            packageID = id;
        }

        public int getPackageID()
        {
            return packageID;
        }
        public void readFluidFromGQIT(int ID)
        {
            try
            {
                thermoSystem = thermoSystem.readObject(ID);
             } catch(System.NullReferenceException e) {
                oldMoleFraction = new double[1];
                thermoSystem = (SystemInterface) new SystemSrkCPAstatoil(298, 10);
                thermoSystem.addComponent("methane", 1.0);
                thermoSystem.createDatabase(true);
            }
            
            oldMoleFraction = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
   
            thermoSystem.init(0);
            thermoSystem.useVolumeCorrection(true);
            thermoSystem.init(1);
            thermoSystem.setNumberOfPhases(1);
            thermoSystem.setMaxNumberOfPhases(3);
        }

        public SystemInterface getThermoSystem()
        {
            return thermoSystem;
        }

        public int getNumberOfComponents()
        {
            return thermoSystem.getPhase(0).getNumberOfComponents();
        }

        public void GetCompoundList(ref object compIds, ref object formulae, ref object names, ref object boilTemps, ref object molwts, ref object casnos)
        {
            names = thermoSystem.getComponentNames();
            formulae = thermoSystem.getCompFormulaes();
            compIds = thermoSystem.getComponentNames();
            casnos = thermoSystem.getCASNumbers();
            boilTemps = thermoSystem.getNormalBoilingPointTemperatures();
            molwts = thermoSystem.getMolecularWeights();
        }


        public String[] getComponentIDs()
        {
            return thermoSystem.getComponentNames();
        }


        public void setTPFraction(double T, double P, double[] x, int activePhaseIndex)
        {
            thermoSystem.removeMoles();
            thermoSystem.setMolarComposition(x);
            //     thermoSystem.setPhaseIndex(0, activePhaseIndex);
            //     thermoSystem.init(0, activePhaseIndex); 
            thermoSystem.init(0, 0);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }

        public void setTPFraction(double T, double P, double[] x)
        {
            thermoSystem.removeMoles();
            thermoSystem.setMolarComposition(x);
            thermoSystem.init(0, 0);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }



        public void setTPFraction(double T, double P, double[] x1, double[] x2)
        {
            thermoSystem.removeMoles();
            thermoSystem.setMolarComposition(x1);
            thermoSystem.init(0);
            thermoSystem.getPhase(0).setMoleFractions(x1);
            thermoSystem.getPhase(1).setMoleFractions(x2);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }

        public void initFlashCalc()
        {
            thermoSystem.setMultiPhaseCheck(true);
            //  thermoSystem.setMaxNumberOfPhases(3);
            //thermoSystem.init(0);
        }

        public void setTPFractionFlash(double T, double P, double[] x)
        {
            thermoSystem.removeMoles();
            thermoSystem.setMolarComposition(x);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
            //thermoSystem.init(0);
        }

        public void endFlashCalc()
        {
          //  thermoSystem.setMaxNumberOfPhases(1);
        }

        public void init(string phase, int initType)
        {
            int phasetype = 0;
            PhaseExist = true;

            if (phase.Equals("Vapor"))
            {
                phasetype = 1;
            }
            else if (phase.Equals("Liquid"))
            {
                phasetype = 0;
            }
            else
            {
                phasetype = 1; // stop here - to check for errors
                string nonHandeledPhase = phase;
            }

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
            {
                PhaseExist = false;
            }
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") || thermoSystem.getPhase(0).getPhaseTypeName().Equals("liquid")) && phase.Equals("Vapor"))
            {
                PhaseExist = false;
            }
            
            //    thermoSystem.setPhaseIndex(0, phaseindex);
            //   thermoSystem.setPhaseType(phaseindex, phasetype); // makes the current phase the first one, need to work with 2 phases
            //   thermoSystem.init(initType, phaseindex);  // init(type, 0)
            thermoSystem.setPhaseType(0, phasetype);
            thermoSystem.init(initType, 0);
        }

        public double[] getFugacityCoefficients(string phase, Boolean doInit = true)
        {
            double factor = 1.0;
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
            {
                factor = 10.0;
            }
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") || thermoSystem.getPhase(0).getPhaseTypeName().Equals("oil")) && phase.Equals("Vapor"))
            {
                factor = 10.0;
            }

            double[] fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < fugacityCoef.Length; i++)
            {
                if (factor > 5.0)
                {
                    fugacityCoef[i] = Math.Exp(1.0 + thermoSystem.getPhase(0).getComponent(i).getLogFugasityCoeffisient());
                }
                else
                {
                    fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getFugasityCoeffisient();
                }

            }

            return fugacityCoef;
        }

        public double[] getLogFugacityCoefficients(string phase, Boolean doInit = true)
        {
            double factor = 1.0;
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
            {
                factor = 10.0;
            }
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") || thermoSystem.getPhase(0).getPhaseTypeName().Equals("oil")) && phase.Equals("Vapor"))
            {
                factor = 10.0;
            }

            double[] fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < fugacityCoef.Length; i++)
            {
                if (factor > 5.0)
                {
                   fugacityCoef[i] = 1.0 + thermoSystem.getPhase(0).getComponent(i).getLogFugasityCoeffisient();
                }
                else
                {
                    fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getLogFugasityCoeffisient();
                }

            }

            return fugacityCoef;
        }


        public void HydrateEquilibriumTemperature()
        {
            thermoSystem.setHydrateCheck(true);
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.hydrateFormationTemperature();
            thermoSystem.setHydrateCheck(false);
        }

        public void TPflash()
        {
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();
            thermoSystem.init(3);
        }

        public void phaseFractionFlash(String phase, double fraction)
        {
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.constantPhaseFractionPressureFlash(fraction);
            thermoSystem.init(3);
        }

        public void PHflash(double enthalpySpec)
        {
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.PHflash(enthalpySpec, 0);
        }

        public void PSflash(double entropySpec)
        {
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.PSflash(entropySpec);
        }

        public double getTemperature()
        {
            return thermoSystem.getTemperature();
        }

        public double getPressure()
        {
            return thermoSystem.getPressure();
        }

        public Boolean checkIfInitNeed(double T, double P, double[] x, string phase)
        {

            int phasetype = 0;
            if (phase.Equals("Vapor"))
            {
                phasetype = 1;
            }


            //if (true)
            //setCurrentProps(T, P, x, phasetype);
            //return true;

            if (oldPhaseType != phasetype)
            {
                setCurrentProps(T, P, x, phasetype);
                return true;
            }

            double sum = 0.0;
            sum += Math.Abs(T - oldTemperature) + Math.Abs(P - oldPressure);
            if (sum > 1e-50)
            {
                setCurrentProps(T, P, x, phasetype);
                return true;
            }

            for (int i = 0; i < x.Length; i++)
            {
                sum += Math.Abs(x[i] - oldMoleFraction[i]);
                if (sum > 1e-50) break;
            }
            if (sum > 1e-50)
            {
                setCurrentProps(T, P, x, phasetype);
                return true;
            }

            return false;
        }

        public void setCurrentProps(double T, double P, double[] x, int localOldPhaseType)
        {
            oldPhaseType = localOldPhaseType;
            oldTemperature = T;
            oldPressure = P;
            Array.Copy(x, oldMoleFraction, x.Length);
        }

        public double[,] getlogFugacityCoefficientsDmoles(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(3);
            }

            double[,] fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents(), thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < fugacityCoef.Length; i++)
            {
                for (int j = 0; j < fugacityCoef.Length; j++)
                {
                    fugacityCoef[i, j] = thermoSystem.getPhase(0).getComponent(i).getdfugdn(j);
                }
            }
            return fugacityCoef;
        }

        public double[] getlogFugacityCoefficientsDtemperature(string phase, Boolean doInit = true)
        {
            //TODO write your implementation code here:
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            double[] fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < fugacityCoef.Length; i++)
            {
                fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getdfugdt();
            }

            return fugacityCoef;
        }




        public double[] getlogFugacityCoefficientsDpressure(string phase, Boolean doInit = true)
        {
            //TODO write your implementation code here:
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            double[] fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < fugacityCoef.Length; i++)
            {
                fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getdfugdp() / 1.0e5;
            }

            return fugacityCoef;
        }

        public double getSurfaceTension(int phase1, int phase2)
        {
            thermoSystem.calcInterfaceProperties();
            return thermoSystem.getInterphaseProperties().getSurfaceTension(phase1, phase2);
        }

        public double getSurfaceTension(string phase1, string phase2)
        {
            string phase1NeqSim = "gas";
            string phase2NeqSim = "liquid";

            if (phase1.Equals("Vapor")) phase1NeqSim = "gas";
            if (phase1.Equals("Liquid")) phase1NeqSim = "liquid";
            if (phase1.Equals("Aqueous")) phase1NeqSim = "aqueous";

            if (phase2.Equals("Vapor")) phase2NeqSim = "gas";
            else if (phase2.Equals("Liquid")) phase2NeqSim = "liquid";
            if (phase2.Equals("Aqueous")) phase2NeqSim = "aqueous";

            int phaseNumb1 = thermoSystem.getPhaseNumberOfPhase(phase1NeqSim);
            int phaseNumb2 = thermoSystem.getPhaseNumberOfPhase(phase2NeqSim);
            int phaseNumb3 = thermoSystem.getPhaseNumberOfPhase("aqueous");

            if (phaseNumb1 == phaseNumb2) phaseNumb2 = phaseNumb3;

            thermoSystem.calcInterfaceProperties();
            return thermoSystem.getInterphaseProperties().getSurfaceTension(phaseNumb1, phaseNumb2);
        }

        public double getEnthalpy(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getEnthalpy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEnthalpydT(string phase, Boolean doInit = true)
        {
            return getHeatCapacityCp(phase, doInit);
        }

        public double getEnthalpydP(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getEnthalpydP() / 1.0e5 / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getInternalEnergydT(string phase, Boolean doInit = true)
        {
            return getHeatCapacityCv(phase, doInit);
        }

        public double[] getEnthalpydN(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            double[] entdn = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < entdn.Length; i++)
            {
                entdn[i] = thermoSystem.getPhase(0).getComponent(i).getEnthalpy(thermoSystem.getPhase(0).getTemperature()) / thermoSystem.getPhase(0).getComponent(i).getNumberOfMolesInPhase();
            }

            return entdn;
        }

        public double[] getEntropydN(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            double[] entdn = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < entdn.Length; i++)
            {
                entdn[i] = thermoSystem.getPhase(0).getComponent(i).getEntropy(thermoSystem.getPhase(0).getTemperature(), thermoSystem.getPhase(0).getPressure()) / thermoSystem.getPhase(0).getComponent(i).getNumberOfMolesInPhase();
            }

            return entdn;
        }




        //TODO write your implementation code here:


        public double getEntropy(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getEntropy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEntropydP(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getEntropydP() / 1.0e5 / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEntropydT(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getEntropydT() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getHelmholtzEnergy(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            return thermoSystem.getPhase(0).getHelmholtzEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getGibbsEnergy(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            return thermoSystem.getPhase(0).getGibbsEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getInternalEnergy(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getInternalEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }



        public double getJouleThomsonCoefficient(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getJouleThomsonCoefficient() / 1.0e5;
        }

        public double getSpeedOfSound(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getSoundSpeed();
        }

        public double getHeatCapacityCp(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getCp() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getHeatCapacityCv(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getCv() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getDensity(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            thermoSystem.getPhase(0).initPhysicalProperties("density");
            //  thermoSystem.useVolumeCorrection(true);
            return thermoSystem.getPhase(0).getPhysicalProperties().getDensity() / thermoSystem.getPhase(0).getMolarMass();
        }

        public double getCompressibilityFactor(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            return thermoSystem.getPhase(0).getZ();
        }

        public double getMolecularWeight(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            return thermoSystem.getPhase(0).getMolarMass() * 1000.0;
        }

        public double getDensitydT(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }
            return thermoSystem.getPhase(0).getdrhodT() / thermoSystem.getPhase(0).getMolarMass();
        }

        public double getDensitydP(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            //thermoSystem.getPhase(0).initPhysicalProperties("density");
            return thermoSystem.getPhase(0).getdrhodP() / 1.0e5 / thermoSystem.getPhase(0).getMolarMass();
        }

        public double[] getDensitydN(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            double[] molMass = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < molMass.Length; i++)
            {
                molMass[i] = thermoSystem.getPhase(0).getComponent(i).getMolarMass();
            }

            return molMass;
        }

        public double[] getMolecularWeightdN(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            double[] molMass = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int i = 0; i < molMass.Length; i++)
            {
                molMass[i] = (thermoSystem.getPhase(0).getComponent(i).getMolarMass() / thermoSystem.getPhase(0).getNumberOfMolesInPhase() - thermoSystem.getPhase(0).getMolarMass() / thermoSystem.getPhase(0).getNumberOfMolesInPhase()) * 1000.0;
            }

            return molMass;
        }

        public double getVolume(string phase, Boolean doInit = true)
        {
            return 1.0 / getDensity(phase, doInit);
            /*
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            return thermoSystem.getPhase(0).getMolarVolume() / 1e5;
             * */
        }

        public double getVolumedT(string phase, Boolean doInit = true)
        {
            double densit2 = getDensity(phase, doInit);
            return -1.0 / (densit2 * densit2) * getDensitydT(phase, doInit);

        }

        public double getVolumedP(string phase, Boolean doInit = true)
        {
            double densit2 = getDensity(phase, doInit);
            return -1.0 / (densit2 * densit2) * getDensitydP(phase, doInit);

        }

        public int[] getOilFractionIDs()
        {
            return thermoSystem.getOilFractionIDs();
        }

        public double[] getOilFractionNormalBoilingPoints()
        {
            return thermoSystem.getOilFractionNormalBoilingPoints();
        }

        public double[] getOilFractionLiquidDensityAt25C()
        {
            return thermoSystem.getOilFractionLiquidDensityAt25C();
        }

        public double[] getOilFractionMolecularMass()
        {
            return thermoSystem.getOilFractionMolecularMass();
        }

        public string[] readCapeOpenProperties10()
        {
            return thermoSystem.getCapeOpenProperties10();
        }

        public string[] readCapeOpenProperties11()
        {
            return thermoSystem.getCapeOpenProperties11();
        }

        public double[] getCriticalVolumes()
        {
            double[] VC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < VC.Length; j++)
            {
                VC[j] = thermoSystem.getPhase(0).getComponent(j).getCriticalVolume() / 1e6;
            }
            return VC;
        }

        public double getCriticalVolume(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getCriticalVolume() / 1e6;
        }

        public double[] getCriticalPressures()
        {
            double[] PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < PC.Length; j++)
            {
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getPC() * 1e5;
            }
            return PC;
        }

        public double getCriticalPressure(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getPC() * 1e5;
        }

        public double[] getCriticalTemperatures()
        {
            double[] TC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < TC.Length; j++)
            {
                TC[j] = thermoSystem.getPhase(0).getComponent(j).getTC();
            }
            return TC;
        }

        public double getCriticalTemperature(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getTC();
        }

        public double[] getAcentricFactor()
        {
            double[] PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < PC.Length; j++)
            {
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getAcentricFactor();
            }
            return PC;
        }


        public double getAcentricFactor(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getAcentricFactor();
        }

        public double[] getNormalBoilingPoint()
        {
            double[] PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < PC.Length; j++)
            {
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getNormalBoilingPoint() + 273.15;
            }
            return PC;
        }

        public double getNormalBoilingPoint(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getNormalBoilingPoint() + 273.15;
        }

        public double[] getMolecularWeight()
        {
            double[] PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < PC.Length; j++)
            {
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getMolarMass() * 1000.0;
            }
            return PC;
        }

        public double getMolecularWeight(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getMolarMass() * 1000.0;
        }

        public double[] getLiquidDensityAt25C()
        {
            double[] PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (int j = 0; j < PC.Length; j++)
            {
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getNormalLiquidDensity() * 1000.0 / thermoSystem.getPhase(0).getComponent(j).getMolarMass();
            }
            return PC;
        }

        public double getLiquidDensityAt25C(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getNormalLiquidDensity() * 1000.0 / thermoSystem.getPhase(0).getComponent(compID).getMolarMass();
        }

        public void addCapeOpenProperty(String propertyName)
        {
            thermoSystem.addCapeOpenProperty(propertyName);
        }

        public double getViscosity(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            thermoSystem.getPhase(0).initPhysicalProperties("viscosity");
            return thermoSystem.getPhase(0).getPhysicalProperties().getViscosity();
        }

        public double getThermalConductivity(string phase, Boolean doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                int phasetype = 0;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }
            thermoSystem.getPhase(0).initPhysicalProperties("conductivity");
            return thermoSystem.getPhase(0).getPhysicalProperties().getConductivity();
        }


        public object getPhaseComposition(int phaseNumber)
        {
            double[] comp = new double[thermoSystem.getPhase(phaseNumber).getNumberOfComponents()];
            for (int i = 0; i < comp.Length; i++)
            {
                comp[i] = thermoSystem.getPhase(phaseNumber).getComponent(i).getx();
            }
            return comp;
        }

        public string[] getPresentPhases()
        {
            string[] phases = new string[thermoSystem.getNumberOfPhases()];

            for (int i = 0; i < phases.Length; i++)
            {
                if (thermoSystem.getPhase(i).getPhaseTypeName().Equals("gas")) phases[i] = "Vapor";
                if (thermoSystem.getPhase(i).getPhaseTypeName().Equals("oil")) phases[i] = "Liquid";
                if (thermoSystem.getPhase(i).getPhaseTypeName().Equals("aqueous")) phases[i] = "Liquid2";
            }
            return phases;

        }

        public double getPhaseFraction(int phaseNumber)
        {
            return thermoSystem.getPhase(phaseNumber).getBeta();
        }

        public double temperature
        {
            get
            {
                return thermoSystem.getTemperature();
            }

        }

        private bool phaseExist;
        public bool PhaseExist
        {
            get
            {
                return this.phaseExist;
            }

            set
            {
                this.phaseExist = value;
            }
        }
    }
}
