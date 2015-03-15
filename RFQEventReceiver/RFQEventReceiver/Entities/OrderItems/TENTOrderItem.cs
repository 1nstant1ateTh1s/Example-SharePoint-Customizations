using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver.Entities.OrderItems
{
    public class TENTOrderItem : OrderItemBase
    {
        [ColumnAttributes("Procurement Agreement Compliant", "Procurement Agreement Compliant")]
        public string ProcurementAgreementCompliant { get; set; }

        [ColumnAttributes("Alternate Procurement Agreement Compliant", "Alternate Procurement Agreement Compliant")]
        public string AlternateProcurementAgreementCompliant { get; set; }
    }
}
