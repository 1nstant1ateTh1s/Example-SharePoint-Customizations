using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace RFQ_SharePoint_Project
{
    [DataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class FES_RFQOrderItem : RFQOrderItem
    {
        // Not serialized - private properties
        private string _tradeAgreementCompliant;
        private string _altTradeAgreementCompliant;

        // Serialized properties        
        [DataMember]
        public string Trade_Agreement_Compliant
        {
            get { return _tradeAgreementCompliant == null ? "" : _tradeAgreementCompliant; }
            set { _tradeAgreementCompliant = value; }
        }
        
        [DataMember(Name = "Alt_Trade_Agreement_Compliant")]
        public string Alternate_Trade_Agreement_Compliant
        {
            get { return _altTradeAgreementCompliant == null ? "" : _altTradeAgreementCompliant; }
            set { _altTradeAgreementCompliant = value; }
        }
    }
}
