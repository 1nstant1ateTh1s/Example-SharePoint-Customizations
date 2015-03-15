using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace RFQ_SharePoint_Project
{
    [DataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class SOE_RFQOrderItem : RFQOrderItem
    {
        // Not serialized - private properties
        private string _procurementAgreementCompliant;
        private string _altProcurementAgreementCompliant;

        // Serialized properties
        [DataMember]
        public string Procurement_Agreement_Compliant
        {
            get { return _procurementAgreementCompliant == null ? "" : _procurementAgreementCompliant; }
            set { _procurementAgreementCompliant = value; }
        }

        [DataMember(Name = "Alt_PA_Compliant")]
        public string Alternate_Procurement_Agreement_Compliant
        {
            get { return _altProcurementAgreementCompliant == null ? "" : _altProcurementAgreementCompliant; }
            set { _altProcurementAgreementCompliant = value; }
        }
    }
}
