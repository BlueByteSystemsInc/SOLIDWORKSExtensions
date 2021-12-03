using System;
using System.Runtime.Serialization;

namespace BlueByte.SOLIDWORKS.Extensions.Helpers
{
    [Serializable]
    public class UnloadAddInsFailedException : Exception
    {
        public UnloadAddInsFailedException()
        {
        }

        public UnloadAddInsFailedException(string message) : base(message)
        {
        }

        public UnloadAddInsFailedException(string message, Exception innerException) : base(message, innerException)
        {
        }


        protected UnloadAddInsFailedException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}