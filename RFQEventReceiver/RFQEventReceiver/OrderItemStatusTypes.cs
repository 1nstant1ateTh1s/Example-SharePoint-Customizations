using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver
{
    public static class OrderItemStatusTypes
    {
        public const string PROCESSING = "Processing";
        public const string AWARDED = "Won";
        public const string LOSS = "Lost";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rfqStatus"></param>
        /// <returns></returns>
        public static string TranslateStatus(string rfqStatus)
        {
            string orderItemStatus = "";
            switch (rfqStatus)
            {
                case RFQStatusTypes.AWARDED:
                    orderItemStatus = AWARDED;
                    break;
                case RFQStatusTypes.LOSS:
                    orderItemStatus = LOSS;
                    break;
                default:
                    orderItemStatus = PROCESSING;
                    break;
            }
            return orderItemStatus;
        }
    }
}
