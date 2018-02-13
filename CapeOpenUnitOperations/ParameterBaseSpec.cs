using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterBaseSpec : ICapeRealParameterSpec, ICapeParameterSpec
    {
        double defaultVal = 1.0;
        double upperBoundVal = 10.0;
        double lowerBoundVal = 0.0;
        CapeParamType paramType = CapeParamType.CAPE_REAL;
        object paramDimensionality = new double[3];
        public Double DefaultValue
        {
            get
            {
                return defaultVal;
            }
        }

        public CapeParamType Type
        {
            get
            {
                return paramType;
            }
             set
            {
                paramType = value;
            }
        }

        public object Dimensionality
        {
            get
            {
                return paramDimensionality;
            }
              set
            {
                paramDimensionality = value;
            }
        }



        public double UpperBound
        {
            get
            {
                return upperBoundVal;
            }
              set
            {
                upperBoundVal = value;
            }
        }

        public double LowerBound
        {
            get
            {
                return lowerBoundVal;
            }
            set
            {
                lowerBoundVal = value;
            }
        }

        public bool Validate(double value, ref string input)
        {
            return true;
        }

    }



    public class ParameterIntSpec : ICapeIntegerParameterSpec, ICapeParameterSpec
    {
        int defaultVal = 1;
        int upperBoundVal = 1000;
        int lowerBoundVal = 0;
        object paramDimensionality = new double[3];
        CapeParamType paramType = CapeParamType.CAPE_INT;

         public int DefaultValue
        {
            get
            {
                return defaultVal;
            }
        }

         public CapeParamType Type
        {
            get
            {
                return paramType;
            }
            set
            {
                paramType = value;
            }
        }

         public object Dimensionality
        {
            get
            {
                return paramDimensionality;
            }
            set
            {
                paramDimensionality = value;
            }
        }



         public int UpperBound
        {
            get
            {
                return upperBoundVal;
            }
            set
            {
                upperBoundVal = value;
            }
        }

         public int LowerBound
        {
            get
            {
                return lowerBoundVal;
            }
            set
            {
                lowerBoundVal = value;
            }
        }

         public bool Validate(int value, ref string input)
        {
            return true;
        }

    }

    public class ParameterRealSpec : ICapeRealParameterSpec, ICapeParameterSpec
    {
        double defaultVal = 1.0;
        double upperBoundVal = 1000.0;
        double lowerBoundVal = 0.0;
        object paramDimensionality = new double[3];
        CapeParamType paramType = CapeParamType.CAPE_REAL;

         public double DefaultValue
        {
            get
            {
                return defaultVal;
            }
        }

         public CapeParamType Type
        {
            get
            {
                return paramType;
            }
            set
            {
                paramType = value;
            }
        }

         public object Dimensionality
        {
            get
            {
                return paramDimensionality;
            }
            set
            {
                paramDimensionality = value;
            }
        }



         public double UpperBound
        {
            get
            {
                return upperBoundVal;
            }
            set
            {
                upperBoundVal = value;
            }
        }

         public double LowerBound
        {
            get
            {
                return lowerBoundVal;
            }
            set
            {
                lowerBoundVal = value;
            }
        }

         public bool Validate(double value, ref string input)
        {
            return true;
        }

    }
}
