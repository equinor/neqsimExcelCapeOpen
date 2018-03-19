using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CAPEOPEN110;

namespace CapeOpenThermo
{
    class PhaseDoesNotExcistExeption : Exception, ECapeThrmPropertyNotAvailable
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
