using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQ_SharePoint_Project
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

        /// <summary>
        /// Determines the variance of Approval status to return based on the quote value.
        /// </summary>
        /// <param name="quoteValue">The total Request for Quote quote value.</param>
        /// <returns>A variance of Approval status type based on the total quote value.</returns>
        public static string GetApprovalStatusType(double quoteValue)
        {
            string approvalStatus = "";

            if (quoteValue >= 1000000.00) // status for quote value greater than or equal to $1M:
            {
                approvalStatus = APPROVAL_GREATERTHAN_1M;
            }
            else if (quoteValue >= 100000.00) // status for quote value greater than or equal to $100K:
            {
                approvalStatus = APPROVAL_GREATERTHAN_100K;
            }
            else // status for quote value less than $100K:
            {
                approvalStatus = APPROVAL_LESSTHAN_100K;
            }

            return approvalStatus; // return the appropriate approval status type
        }
    }
}
