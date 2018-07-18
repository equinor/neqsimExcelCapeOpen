using CAPEOPEN110;

namespace CapeOpenUnitOperations
{
    public class ParameterBaseSpec : ICapeRealParameterSpec, ICapeParameterSpec
    {
        public CapeParamType Type { get; set; } = CapeParamType.CAPE_REAL;

        public object Dimensionality { get; set; } = new double[3];

        public double DefaultValue { get; } = 1.0;


        public double UpperBound { get; set; } = 10.0;

        public double LowerBound { get; set; } = 0.0;

        public bool Validate(double value, ref string input)
        {
            return true;
        }
    }


    public class ParameterIntSpec : ICapeIntegerParameterSpec, ICapeParameterSpec
    {
        public int DefaultValue { get; } = 1;


        public int UpperBound { get; set; } = 1000;

        public int LowerBound { get; set; } = 0;

        public bool Validate(int value, ref string input)
        {
            return true;
        }

        public CapeParamType Type { get; set; } = CapeParamType.CAPE_INT;

        public object Dimensionality { get; set; } = new double[3];
    }

    public class ParameterRealSpec : ICapeRealParameterSpec, ICapeParameterSpec
    {
        public CapeParamType Type { get; set; } = CapeParamType.CAPE_REAL;

        public object Dimensionality { get; set; } = new double[3];

        public double DefaultValue { get; } = 1.0;


        public double UpperBound { get; set; } = 1000.0;

        public double LowerBound { get; set; } = 0.0;

        public bool Validate(double value, ref string input)
        {
            return true;
        }
    }
}