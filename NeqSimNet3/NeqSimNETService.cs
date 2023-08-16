using System;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using org.apache.logging.log4j.core.appender;
using org.apache.logging.log4j.core.layout;

namespace NeqSimNET
{
    public class NeqSimNETService
    {
        public static int packageID;

        private double[] oldMoleFraction;
        private int oldPhaseType = 2;
        private double oldPressure = 1.00112;
        private double oldTemperature = 1.00112;

        // public bool PhaseExist { get => PhaseExist; set => PhaseExist = value; }

        private SystemInterface thermoSystem = new SystemSrkEos(298.0, 10.0);

        public double temperature => thermoSystem.getTemperature();

        public string[] CapeOpenProperties { get; set; }

        public bool PhaseExist { get; set; }

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
            }
            catch (NullReferenceException e)
            {
                oldMoleFraction = new double[1];
                thermoSystem = new SystemSrkEos(298, 10);
                thermoSystem.addComponent("methane", 1.0);
                thermoSystem.createDatabase(true);
                thermoSystem.setMixingRule(2);
                System.Diagnostics.Debug.WriteLine(e.Message);
            }


            thermoSystem.init(0);
            thermoSystem.setMultiPhaseCheck(true);
            thermoSystem.useVolumeCorrection(true);
            thermoSystem.init(1);
            thermoSystem.setNumberOfPhases(1);
            thermoSystem.setMaxNumberOfPhases(3);
            oldMoleFraction = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
        }

        public void readFluidFromGQIT(string ID)
        {
            try
            {
                thermoSystem = thermoSystem.readObjectFromFile(ID, ID);
            }
            catch (NullReferenceException e)
            {
                oldMoleFraction = new double[1];
                thermoSystem = new SystemSrkEos(298, 10);
                thermoSystem.addComponent("methane", 1.0);
                thermoSystem.createDatabase(true);
                thermoSystem.setMixingRule(2);
                System.Diagnostics.Debug.WriteLine(e.Message);
            }

            oldMoleFraction = new double[thermoSystem.getPhase(0).getNumberOfComponents()];

            thermoSystem.init(0);
            thermoSystem.setMultiPhaseCheck(true);
            thermoSystem.useVolumeCorrection(true);
            thermoSystem.init(1);
            thermoSystem.setNumberOfPhases(1);
            thermoSystem.setMaxNumberOfPhases(3);
            oldMoleFraction = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
        }

        public SystemInterface getThermoSystem()
        {
            return thermoSystem;
        }

        public int getNumberOfComponents()
        {
            return thermoSystem.getPhase(0).getNumberOfComponents();
        }

        public void GetCompoundList(ref object compIds, ref object formulae, ref object names, ref object boilTemps,
            ref object molwts, ref object casnos)
        {
            names = thermoSystem.getComponentNames();
            formulae = thermoSystem.getCompFormulaes();
            compIds = thermoSystem.getComponentNames();
            casnos = thermoSystem.getCASNumbers();
            boilTemps = thermoSystem.getNormalBoilingPointTemperatures();
            molwts = thermoSystem.getMolecularWeights();
        }


        public string[] getComponentIDs()
        {
            return thermoSystem.getComponentNames();
        }


        public void setTPFraction(double T, double P, double[] x, int activePhaseIndex)
        {
            thermoSystem.setMolarComposition(x);
            //     thermoSystem.setPhaseIndex(0, activePhaseIndex);
            //     thermoSystem.init(0, activePhaseIndex);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }

        public void setTPFraction(double T, double P, double[] x)
        {
          
            thermoSystem.init(0);
            for (int i = 0; i < x.Length; i++) {
                if (x[i] < 0) x[i] = 0.0;
            }

            
            thermoSystem.setMolarComposition(x);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
            thermoSystem.setNumberOfPhases(1);
        }


        public void setTPFraction(double T, double P, double[] x1, double[] x2)
        {
            //thermoSystem.removeMoles();
            thermoSystem.setMolarComposition(x1);
            thermoSystem.getPhase(0).setMoleFractions(x1);
            thermoSystem.getPhase(1).setMoleFractions(x2);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }

        public void initFlashCalc()
        {
            // thermoSystem.setMultiPhaseCheck(true);
        }

        public void setTPFractionFlash(double T, double P, double[] x)
        {
            setCurrentProps(T, P, x, 0);
            thermoSystem.setMolarComposition(x);
            thermoSystem.setTemperature(T);
            thermoSystem.setPressure(P);
        }

        public void endFlashCalc()
        {
            //  thermoSystem.setMaxNumberOfPhases(1);
        }

        public void init(string phase, int initType)
        {

            var phasetype = 0;
            var phaseIndex = 1;
           
            //  PhaseExist = true;

            if (phase.Equals("Vapor"))
            {
                phasetype = 1;
                phaseIndex = 0;
            }
            else if (phase.Equals("Liquid"))
            {
                phasetype = 0;
                phaseIndex = 1;
            }
            else
            {
                phasetype = 0; // stop here - to check for errors
                phaseIndex = 0;
                var nonHandeledPhase = phase;
            }

            thermoSystem.setPhaseIndex(0, phaseIndex);
            thermoSystem.setPhaseType(0, phasetype);

            thermoSystem.init(initType, 0);

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
                PhaseExist = false;
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") ||
                      thermoSystem.getPhase(0).getPhaseTypeName().Equals("oil")) &&
                     phase.Equals("Vapor")) PhaseExist = false;
        }

        public double[] getFugacityCoefficients(string phase, bool doInit = true)
        {
            var factor = 1.0;
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }
                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
                factor = 1.0;
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") ||
                      thermoSystem.getPhase(0).getPhaseTypeName().Equals("oil")) &&
                     phase.Equals("Vapor")) factor = 1.0;

            var fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < fugacityCoef.Length; i++)
                if (factor > 5.0)
                    fugacityCoef[i] =
                        Math.Exp(1.0 + thermoSystem.getPhase(0).getComponent(i).getLogFugacityCoefficient());
                else
                    fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getFugacityCoefficient();

            return fugacityCoef;
        }

        public double[] getLogFugacityCoefficients(string phase, bool doInit = true)
        {
            var factor = 1.0;
            if (doInit)
            {
                 thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }
                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            if (thermoSystem.getPhase(0).getPhaseTypeName().Equals("gas") && phase.Equals("Liquid"))
                factor = 1.0;
            else if ((thermoSystem.getPhase(0).getPhaseTypeName().Equals("aqueous") ||
                      thermoSystem.getPhase(0).getPhaseTypeName().Equals("oil")) &&
                     phase.Equals("Vapor")) factor = 1.0;

            var fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < fugacityCoef.Length; i++)
                if (factor > 5.0)
                    fugacityCoef[i] = 1.0 + thermoSystem.getPhase(0).getComponent(i).getLogFugacityCoefficient();
                else
                    fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getLogFugacityCoefficient();

            return fugacityCoef;
        }


        public void HydrateEquilibriumTemperature()
        {
            thermoSystem.setHydrateCheck(true);
            var ops = new ThermodynamicOperations(thermoSystem);
            ops.hydrateFormationTemperature();
            thermoSystem.setHydrateCheck(false);
        }

        public void TPflash()
        {
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();
            //thermoSystem.init(3);
        }

        public void phaseFractionFlash(string phase, double fraction)
        {
            var ops = new ThermodynamicOperations(thermoSystem);
            ops.constantPhaseFractionPressureFlash(fraction);
            thermoSystem.init(3);
        }

        public void PHflash(double enthalpySpec)
        {
            var ops = new ThermodynamicOperations(thermoSystem);
            ops.PHflash(enthalpySpec, "J/mol");
        }

        public void PSflash(double entropySpec)
        {
            var ops = new ThermodynamicOperations(thermoSystem);
            ops.PSflash(entropySpec, "J/molK");
        }

        public double getTemperature()
        {
            return thermoSystem.getTemperature();
        }

        public double getPressure()
        {
            return thermoSystem.getPressure();
        }

        public bool checkIfInitNeed(double T, double P, double[] x, string phase)
        {
            var phasetype = 0;
            if (phase.Equals("Vapor")) phasetype = 1;


            //if (true)
            //setCurrentProps(T, P, x, phasetype);
            //return true;

            if (oldPhaseType != phasetype)
            {
                setCurrentProps(T, P, x, phasetype);
                return true;
            }

            var sum = 0.0;
            sum += Math.Abs(T - oldTemperature) + Math.Abs(P - oldPressure);
            if (sum > 1e-50)
            {
                setCurrentProps(T, P, x, phasetype);
                return true;
            }

            for (var i = 0; i < x.Length; i++)
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

        public double[,] getlogFugacityCoefficientsDmoles(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(3);
            }

            var fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents(),
                thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < fugacityCoef.Length; i++)
                for (var j = 0; j < fugacityCoef.Length; j++)
                    fugacityCoef[i, j] = thermoSystem.getPhase(0).getComponent(i).getdfugdn(j);
            return fugacityCoef;
        }

        public double[] getlogFugacityCoefficientsDtemperature(string phase, bool doInit = true)
        {
            //TODO write your implementation code here:
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            var fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < fugacityCoef.Length; i++)
                fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getdfugdt();

            return fugacityCoef;
        }


        public double[] getlogFugacityCoefficientsDpressure(string phase, bool doInit = true)
        {
            //TODO write your implementation code here:
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            var fugacityCoef = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < fugacityCoef.Length; i++)
                fugacityCoef[i] = thermoSystem.getPhase(0).getComponent(i).getdfugdp() / 1.0e5;

            return fugacityCoef;
        }

        public double getSurfaceTension(int phase1, int phase2)
        {
            thermoSystem.calcInterfaceProperties();
            return thermoSystem.getInterphaseProperties().getSurfaceTension(phase1, phase2);
        }

        public double getSurfaceTension(string phase1, string phase2)
        {
            var phase1NeqSim = "gas";
            var phase2NeqSim = "liquid";

            if (phase1.Equals("Vapor")) phase1NeqSim = "gas";
            if (phase1.Equals("Liquid")) phase1NeqSim = "liquid";
            if (phase1.Equals("Aqueous")) phase1NeqSim = "aqueous";

            if (phase2.Equals("Vapor")) phase2NeqSim = "gas";
            else if (phase2.Equals("Liquid")) phase2NeqSim = "liquid";
            if (phase2.Equals("Aqueous")) phase2NeqSim = "aqueous";

            var phaseNumb1 = thermoSystem.getPhaseNumberOfPhase(phase1NeqSim);
            var phaseNumb2 = thermoSystem.getPhaseNumberOfPhase(phase2NeqSim);
            var phaseNumb3 = thermoSystem.getPhaseNumberOfPhase("aqueous");

            if (phaseNumb1 == phaseNumb2) phaseNumb2 = phaseNumb3;

            thermoSystem.calcInterfaceProperties();
            return thermoSystem.getInterphaseProperties().getSurfaceTension(phaseNumb1, phaseNumb2);
        }

        public double getEnthalpy(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getEnthalpy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEnthalpydT(string phase, bool doInit = true)
        {
            return getHeatCapacityCp(phase, doInit);
        }

        public double getEnthalpydP(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getEnthalpydP() / 1.0e5 /
                   thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getInternalEnergydT(string phase, bool doInit = true)
        {
            return getHeatCapacityCv(phase, doInit);
        }

        public double[] getEnthalpydN(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            var entdn = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < entdn.Length; i++)
                entdn[i] = thermoSystem.getPhase(0).getComponent(i)
                               .getEnthalpy(thermoSystem.getPhase(0).getTemperature()) / thermoSystem.getPhase(0)
                               .getComponent(i).getNumberOfMolesInPhase();

            return entdn;
        }

        public double[] getEntropydN(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.init(2);
            }

            var entdn = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < entdn.Length; i++)
                entdn[i] = thermoSystem.getPhase(0).getComponent(i)
                               .getEntropy(thermoSystem.getPhase(0).getTemperature(),
                                   thermoSystem.getPhase(0).getPressure()) / thermoSystem.getPhase(0).getComponent(i)
                               .getNumberOfMolesInPhase();

            return entdn;
        }


        //TODO write your implementation code here:


        public double getEntropy(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getEntropy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEntropydP(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getEntropydP() / 1.0e5 / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getEntropydT(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getEntropydT() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getHelmholtzEnergy(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            return thermoSystem.getPhase(0).getHelmholtzEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getGibbsEnergy(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            return thermoSystem.getPhase(0).getGibbsEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getInternalEnergy(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getInternalEnergy() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }


        public double getJouleThomsonCoefficient(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getJouleThomsonCoefficient() / 1.0e5;
        }

        public double getSpeedOfSound(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getSoundSpeed();
        }

        public double getHeatCapacityCp(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getCp() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getHeatCapacityCv(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getCv() / thermoSystem.getPhase(0).getNumberOfMolesInPhase();
        }

        public double getDensity(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            thermoSystem.getPhase(0).initPhysicalProperties("density");
            //   thermoSystem.useVolumeCorrection(false);
            return thermoSystem.getPhase(0).getPhysicalProperties().getDensity() /
                   thermoSystem.getPhase(0).getMolarMass();
        }

        public double getCompressibilityFactor(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            return thermoSystem.getPhase(0).getZ();
        }

        public double getMolecularWeight(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            return thermoSystem.getPhase(0).getMolarMass() * 1000.0;
        }

        public double getDensitydT(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(2);
            }

            return thermoSystem.getPhase(0).getdrhodT() / thermoSystem.getPhase(0).getMolarMass();
        }

        public double getDensitydP(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            //thermoSystem.getPhase(0).initPhysicalProperties("density");
            return thermoSystem.getPhase(0).getdrhodP() / 1.0e5 / thermoSystem.getPhase(0).getMolarMass();
        }

        public double[] getDensitydN(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }
                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            var molMass = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < molMass.Length; i++)
                molMass[i] = thermoSystem.getPhase(0).getComponent(i).getMolarMass();

            return molMass;
        }

        public double[] getMolecularWeightdN(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }
                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            var molMass = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var i = 0; i < molMass.Length; i++)
                molMass[i] =
                    (thermoSystem.getPhase(0).getComponent(i).getMolarMass() /
                     thermoSystem.getPhase(0).getNumberOfMolesInPhase() - thermoSystem.getPhase(0).getMolarMass() /
                     thermoSystem.getPhase(0).getNumberOfMolesInPhase()) * 1000.0;

            return molMass;
        }

        public double getVolume(string phase, bool doInit = true)
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

        public double getVolumedT(string phase, bool doInit = true)
        {
            var densit2 = getDensity(phase, doInit);
            return -1.0 / (densit2 * densit2) * getDensitydT(phase, doInit);
        }

        public double getVolumedP(string phase, bool doInit = true)
        {
            var densit2 = getDensity(phase, doInit);
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
            CapeOpenProperties = thermoSystem.getCapeOpenProperties10();
            return thermoSystem.getCapeOpenProperties10();
        }

        public string[] readCapeOpenProperties11()
        {
            CapeOpenProperties = thermoSystem.getCapeOpenProperties11();
            return thermoSystem.getCapeOpenProperties11();
        }

        public double[] getCriticalVolumes()
        {
            var VC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < VC.Length; j++)
                VC[j] = thermoSystem.getPhase(0).getComponent(j).getCriticalVolume() / 1e6;
            return VC;
        }

        public double getCriticalVolume(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getCriticalVolume() / 1e6;
        }

        public double[] getCriticalPressures()
        {
            var PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < PC.Length; j++) PC[j] = thermoSystem.getPhase(0).getComponent(j).getPC() * 1e5;
            return PC;
        }

        public double getCriticalPressure(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getPC() * 1e5;
        }

        public double[] getCriticalTemperatures()
        {
            var TC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < TC.Length; j++) TC[j] = thermoSystem.getPhase(0).getComponent(j).getTC();
            return TC;
        }

        public double getCriticalTemperature(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getTC();
        }

        public double[] getAcentricFactor()
        {
            var PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < PC.Length; j++) PC[j] = thermoSystem.getPhase(0).getComponent(j).getAcentricFactor();
            return PC;
        }


        public double getAcentricFactor(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getAcentricFactor();
        }

        public double[] getNormalBoilingPoint()
        {
            var PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < PC.Length; j++)
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getNormalBoilingPoint() + 273.15;
            return PC;
        }

        public double getNormalBoilingPoint(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getNormalBoilingPoint() + 273.15;
        }

        public double[] getMolecularWeight()
        {
            var PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < PC.Length; j++)
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getMolarMass() * 1000.0;
            return PC;
        }

        public double getMolecularWeight(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getMolarMass() * 1000.0;
        }

        public double[] getLiquidDensityAt25C()
        {
            var PC = new double[thermoSystem.getPhase(0).getNumberOfComponents()];
            for (var j = 0; j < PC.Length; j++)
                PC[j] = thermoSystem.getPhase(0).getComponent(j).getNormalLiquidDensity() * 1000.0 /
                        thermoSystem.getPhase(0).getComponent(j).getMolarMass();
            return PC;
        }

        public double getLiquidDensityAt25C(string compID)
        {
            return thermoSystem.getPhase(0).getComponent(compID).getNormalLiquidDensity() * 1000.0 /
                   thermoSystem.getPhase(0).getComponent(compID).getMolarMass();
        }

        public void addCapeOpenProperty(string propertyName)
        {
            thermoSystem.addCapeOpenProperty(propertyName);
        }

        public double getViscosity(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            thermoSystem.getPhase(0).initPhysicalProperties("viscosity");
            return thermoSystem.getPhase(0).getPhysicalProperties().getViscosity();
        }

        public double getThermalConductivity(string phase, bool doInit = true)
        {
            if (doInit)
            {
                thermoSystem.setNumberOfPhases(1);
                var phasetype = 0;
                var phaseIndex = 1;
                if (phase.Equals("Vapor"))
                {
                    phasetype = 1;
                    phaseIndex = 0;
                }

                thermoSystem.setPhaseIndex(0, phaseIndex);
                thermoSystem.setPhaseType(0, phasetype);
                thermoSystem.init(1);
            }

            thermoSystem.getPhase(0).initPhysicalProperties("conductivity");
            return thermoSystem.getPhase(0).getPhysicalProperties().getConductivity();
        }


        public object getPhaseComposition(int phaseNumber)
        {
            var comp = new double[thermoSystem.getPhase(phaseNumber).getNumberOfComponents()];
            for (var i = 0; i < comp.Length; i++) comp[i] = thermoSystem.getPhase(phaseNumber).getComponent(i).getx();
            return comp;
        }

        public string[] getPresentPhases()
        {
            var phases = new string[thermoSystem.getNumberOfPhases()];

            for (var i = 0; i < phases.Length; i++)
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

        public void addComponent(string componentName, double numberOfMoles)
        {
            thermoSystem.addComponent(componentName, numberOfMoles);
        }

        public void setMixingRule(int i)
        {
            thermoSystem.setMixingRule(i);
        }

        public void initFluid(int i)
        {
            thermoSystem.init(i);
        }

        public void createFluid() {
        thermoSystem = new SystemSrkEos(298.0, 10.0);
    }

    }
}