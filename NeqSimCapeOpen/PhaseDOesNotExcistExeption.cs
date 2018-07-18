using System;
using CAPEOPEN110;

namespace CapeOpenThermo
{
    internal class PhaseDoesNotExcistExeption : Exception, ECapeThrmPropertyNotAvailable
    {
        public PhaseDoesNotExcistExeption()
        {
        }

        public PhaseDoesNotExcistExeption(string message)
            : base(message)
        {
        }

        public PhaseDoesNotExcistExeption(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}