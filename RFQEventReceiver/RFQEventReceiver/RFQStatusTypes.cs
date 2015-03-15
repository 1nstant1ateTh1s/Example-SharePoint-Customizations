using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver
{
    public static class RFQStatusTypes
    {
        public const string RECEIVED = "Received";
        public const string UNASSIGNED = "Unassigned";
        public const string APPROVAL_LESSTHAN_100K = "Need Approval<$100k";
        public const string APPROVAL_GREATERTHAN_100K = "Need Approval>$100k";
        public const string APPROVAL_GREATERTHAN_1M = "Need Approval>$1M";
        public const string APPROVED = "Approved";
        public const string SUBMITTED = "Submitted";
        public const string AWARDED = "Awarded";
        public const string LOSS = "Loss";

    }
}
