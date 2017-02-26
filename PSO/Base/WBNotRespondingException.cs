using System;

namespace Iren.PSO.Base
{
    public class WBNotRespondingException : Exception
    {
        public WBNotRespondingException()
        {
        }

        public WBNotRespondingException(string message)
            : base(message)
        {
        }

        public WBNotRespondingException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
