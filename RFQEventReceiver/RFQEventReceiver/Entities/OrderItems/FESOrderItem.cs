using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver.Entities.OrderItems
{
    public class FESOrderItem : OrderItemBase
    {
        [ColumnAttributes("ShipTo DoDAAC", "ShipTo DoDAAC")]
        public override string ShipToDODAAC { get; set; }

        [ColumnAttributes("Trade Agreement Compliant", "Trade Agreement Compliant")]
        public string TradeAgreementCompliant { get; set; }

        [ColumnAttributes("Alternate Trade Agreement Compliant", "Alternate Trade Agreement Compliant")]
        public string AlternateTradeAgreementCompliant { get; set; }
    }
}
