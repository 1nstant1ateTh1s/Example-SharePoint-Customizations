using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver
{
    public class RFQException : Exception
    {
        // Properties
        public string RFQNumber { get; set; }

        // Constructors
        public RFQException(string message, string rfqNumber) : base(message)
        {
            this.RFQNumber = rfqNumber;
        }

        public RFQException(string message, Exception innerException, string rfqNumber)
            : base(message, innerException)
        {
            this.RFQNumber = rfqNumber;
        }
    }
}
