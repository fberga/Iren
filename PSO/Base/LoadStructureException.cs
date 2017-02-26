using System;

namespace Iren.PSO.Base
{
    public class LoadStructureException : Exception
    {
        public LoadStructureException()
        {
        }

        public LoadStructureException(string message)
            : base(message)
        {
        }

        public LoadStructureException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
