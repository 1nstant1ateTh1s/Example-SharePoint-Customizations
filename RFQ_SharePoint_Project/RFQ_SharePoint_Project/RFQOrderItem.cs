using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace RFQ_SharePoint_Project
{
    [DataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class RFQOrderItem
    {
        // Not serialized - private properties
        private string _orderItemId;
        private string _quoteNumber;
        private string _tlspVendorPartNumber;
        private double _transportationPrice;
        private double _vendorUnitPrice;
        private double _purchaseExtendedPrice;
        private string _leadTime;
        private string _comments;
        private string _berryAmendmentCompliant;
        private string _countryofOrigin;
        private string _altCoreListNumber;
        private string _altNSN;
        private string _altManufacturerName;
        private string _altManufacturerPartNumber;
        private string _altTLSPVendorPartNumber;
        private string _altItemDescription;
        private double _altTransportationPrice;
        private double _altVendorUnitPrice;
        private string _altLeadTime;
        private string _altComments;
        private string _altBerryAmendmentCompliant;
        private string _altCountryofOrigin;

        // Serialized properties
        [DataMember(Name = "Title")]
        public string Order_Item_ID
        {
            get { return _orderItemId == null ? "" : _orderItemId; }
            set { _orderItemId = value; }
        }

        [DataMember]
        public string Quote_Number
        {
            get { return _quoteNumber == null ? "" : _quoteNumber; }
            set { _quoteNumber = value; }
        }

        [DataMember]
        public string TLSP_Vendor_Part_Number
        {
            get { return (_tlspVendorPartNumber == null ? "" : _tlspVendorPartNumber); }
            set { _tlspVendorPartNumber = value; }
        }

        [DataMember]
        public double Transportation_Price
        {
            get { return _transportationPrice; }
            set { _transportationPrice = value; }
        }

        [DataMember]
        public double Vendor_Unit_Price
        {
            get { return _vendorUnitPrice; }
            set { _vendorUnitPrice = value; }
        }

        [DataMember]
        public double Purchase_Extended_Price
        {
            get { return _purchaseExtendedPrice; }
            set { _purchaseExtendedPrice = value; }
        }

        [DataMember]
        public string Lead_Time
        {
            get { return _leadTime == null ? "" : _leadTime; }
            set { _leadTime = value; }
        }

        [DataMember]
        public string Comments
        {
            get { return _comments == null ? "" : _comments; }
            set { _comments = value; }
        }

        [DataMember]
        public string Berry_Amendment_Compliant
        {
            get { return _berryAmendmentCompliant == null ? "" : _berryAmendmentCompliant; }
            set { _berryAmendmentCompliant = value; }
        }

        [DataMember]
        public string Country_of_Origin
        {
            get { return _countryofOrigin == null ? "" : _countryofOrigin; }
            set { _countryofOrigin = value; }
        }

        [DataMember(Name = "Alt_Core_List_Number")]
        public string Alternate_Core_List_Number
        {
            get { return _altCoreListNumber == null ? "" : _altCoreListNumber; }
            set { _altCoreListNumber = value; }
        }

        [DataMember(Name = "Alt_NSN")]
        public string Alternate_NSN
        {
            get { return _altNSN == null ? "" : _altNSN; }
            set { _altNSN = value; }
        }

        [DataMember(Name = "Alt_Manufacturer_Name")]
        public string Alternate_Manufacturer_Name
        {
            get { return _altManufacturerName == null ? "" : _altManufacturerName; }
            set { _altManufacturerName = value; }
        }

        [DataMember(Name = "Alt_Manufacturer_Part_Number")]
        public string Alternate_Manufacturer_Part_Number
        {
            get { return _altManufacturerPartNumber == null ? "" : _altManufacturerPartNumber; }
            set { _altManufacturerPartNumber = value; }
        }

        [DataMember(Name = "Alt_TLSP_Vendor_Part_Number")]
        public string Alternate_TLSP_Vendor_Part_Number
        {
            get { return _altTLSPVendorPartNumber == null ? "" : _altTLSPVendorPartNumber; }
            set { _altTLSPVendorPartNumber = value; }
        }

        [DataMember(Name = "Alt_Item_Description")]
        public string Alternate_Item_Description
        {
            get { return _altItemDescription == null ? "" : _altItemDescription; }
            set { _altItemDescription = value; }
        }

        [DataMember(Name = "Alt_Transportation_Price")]
        public double Alternate_Transportation_Price
        {
            get { return _altTransportationPrice; }
            set { _altTransportationPrice = value; }
        }

        [DataMember(Name = "Alt_Vendor_Unit_Price")]
        public double Alternate_Vendor_Unit_Price
        {
            get { return _altVendorUnitPrice; }
            set { _altVendorUnitPrice = value; }
        }

        [DataMember(Name = "Alt_Lead_Time")]
        public string Alternate_Lead_Time
        {
            get { return _altLeadTime == null ? "" : _altLeadTime; }
            set { _altLeadTime = value; }
        }

        [DataMember(Name = "Alt_Comments")]
        public string Alternate_Comments
        {
            get { return _altComments == null ? "" : _altComments; }
            set { _altComments = value; }
        }

        [DataMember(Name = "Alt_Berry_Amendment_Compliant")]
        public string Alternate_Berry_Amendment_Compliant
        {
            get { return _altBerryAmendmentCompliant == null ? "" : _altBerryAmendmentCompliant; }
            set { _altBerryAmendmentCompliant = value; }
        }

        [DataMember(Name = "Alt_Country_of_Origin")]
        public string Alternate_Country_of_Origin
        {
            get { return _altCountryofOrigin == null ? "" : _altCountryofOrigin; }
            set { _altCountryofOrigin = value; }
        }
    }
}
