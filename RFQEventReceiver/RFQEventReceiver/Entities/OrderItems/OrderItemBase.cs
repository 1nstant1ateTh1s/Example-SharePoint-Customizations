using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver.Entities.OrderItems
{
    public class OrderItemBase : ColumnInfo
    {
        [ColumnAttributes("Order Item ID", "Order Item ID")]
        public string OrderItemID { get; set; }

        [ColumnAttributes("Quote Number", "Quote Number")]
        public string RFQQuoteNumber { get; set; }

        [ColumnAttributes("Vendor ID", "Vendor ID")]
        public string VendorID { get; set; }

        [ColumnAttributes("TLSP Vendor Extension", "TLSP Vendor Extension")]
        public string TLSPVendorExtension { get; set; }

        [ColumnAttributes("Region Group", "Region Group")]
        public string RegionGroup { get; set; }

        [ColumnAttributes("Region", "Region")]
        public string Region { get; set; }

        [ColumnAttributes("Request Type", "Request Type")]
        public string RequestType { get; set; }

        [ColumnAttributes("Core List Number", "Core List Number")]
        public string CoreListNumber { get; set; }

        [ColumnAttributes("NSN", "NSN")]
        public string NSN { get; set; }

        [ColumnAttributes("Manufacturer Name", "Manufacturer Name")]
        public string ManufacturerName { get; set; }

        [ColumnAttributes("Manufacturer Part Number", "Manufacturer Part Number")]
        public string ManufacturerPartNumber { get; set; }

        [ColumnAttributes("TLSP Vendor Part Number", "TLSP Vendor Part Number")]
        public string TLSPVendorPartNumber { get; set; }

        [ColumnAttributes("Item Description", "Item Description")]
        public string ItemDescription { get; set; }

        [ColumnAttributes("Additional Information", "Additional Information")]
        public string AdditionalInformation { get; set; }

        [ColumnAttributes("Requested Delivery Date", "Requested Delivery Date")]
        public string RequestedDeliveryDate { get; set; }

        [ColumnAttributes("Unit of Issue", "Unit of Issue")]
        public string UnitOfIssue { get; set; }

        [ColumnAttributes("Quantity", "Quantity")]
        public string Quantity { get; set; }

        [ColumnAttributes("ShipTo DODAAC", "ShipTo DODAAC")]
        public virtual string ShipToDODAAC { get; set; }

        [ColumnAttributes("Delivery Destination", "Delivery Destination")]
        public string DeliveryDestination { get; set; }

        [ColumnAttributes("FOB Origin", "FOB Origin")]
        public string FOBOrigin { get; set; }

        [ColumnAttributes("RFID Required", "RFID Required")]
        public string RFIDRequired { get; set; }

        [ColumnAttributes("Special Project Code", "Special Project Code")]
        public string SpecialProjectCode { get; set; }

        [ColumnAttributes("Transportation Price Required", "Transportation Price Required")]
        public string TransportationPriceRequired { get; set; }

        [ColumnAttributes("Transportation Price", "Transportation Price")]
        public string TransportationPrice { get; set; }

        [ColumnAttributes("Vendor Unit Price", "Vendor Unit Price")]
        public string VendorUnitPrice { get; set; }

        [ColumnAttributes("Purchase Unit Price", "Purchase Unit Price")]
        public string PurchaseUnitPrice { get; set; }

        [ColumnAttributes("Purchase Extended Price", "Purchase Extended Price")]
        public string PurchaseExtendedPrice { get; set; }

        [ColumnAttributes("Lead Time", "Lead Time")]
        public string LeadTime { get; set; }

        [ColumnAttributes("Comments", "Comments")]
        public string Comments { get; set; }

        [ColumnAttributes("Berry Amendment Compliant", "Berry Amendment Compliant")]
        public string BerryAmendmentCompliant { get; set; }

        [ColumnAttributes("Country of Origin", "Country of Origin")]
        public string CountryOfOrigin { get; set; }

        [ColumnAttributes("Customer Permits Alternates", "Customer Permits Alternates")]
        public string CustomerPermitsAlternates { get; set; }

        [ColumnAttributes("Alternate Core List Number", "Alternate Core List Number")]
        public string AltCoreListNumber { get; set; }

        [ColumnAttributes("Alternate NSN", "Alternate NSN")]
        public string AltNSN { get; set; }

        [ColumnAttributes("Alternate Manufacturer Name", "Alternate Manufacturer Name")]
        public string AltManufacturerName { get; set; }

        [ColumnAttributes("Alternate Manufacturer Part Number", "Alternate Manufacturer Part Number")]
        public string AltManufacturerPartNumber { get; set; }

        [ColumnAttributes("Alternate TLSP Vendor Part Number", "Alternate TLSP Vendor Part Number")]
        public string AltTLSPVendorPartNumber { get; set; }

        [ColumnAttributes("Alternate Item Description", "Alternate Item Description")]
        public string AltItemDescription { get; set; }

        [ColumnAttributes("Alternate Transportation Price", "Alternate Transportation Price")]
        public string AltTransportationPrice { get; set; }

        [ColumnAttributes("Alternate Vendor Unit Price", "Alternate Vendor Unit Price")]
        public string AltVendorUnitPrice { get; set; }

        [ColumnAttributes("Alternate Purchase Unit Price", "Alternate Purchase Unit Price")]
        public string AltPurchaseUnitPrice { get; set; }

        [ColumnAttributes("Alternate Purchase Extended Price", "Alternate Purchase Extended Price")]
        public string AltPurchaseExtendedPrice { get; set; }

        [ColumnAttributes("Alternate Lead Time", "Alternate Lead Time")]
        public string AltLeadTime { get; set; }

        [ColumnAttributes("Alternate Comments", "Alternate Comments")]
        public string AltComments { get; set; }

        [ColumnAttributes("Alternate Berry Amendment Compliant", "Alternate Berry Amendment Compliant")]
        public string AltBerryAmendmentCompliant { get; set; }

        [ColumnAttributes("Alternare Country of Origin", "Alternate Country of Origin")]
        public string AltCountryOfOrigin { get; set; }

        [ColumnAttributes("Quote Due Date", "Quote Due Date")]
        public string QuoteDueDate { get; set; }

        [ColumnAttributes("Series Number", null)]
        public int SeriesNumber { get; set; }

        [ColumnAttributes("Series Total", null)]
        public int SeriesTotal { get; set; }

        [ColumnAttributes("Fiscal Year", null)]
        public string FiscalYear {
            get {
                return DetermineRFQFiscalYear();
            }
        }

        [ColumnAttributes("Burdened Unit Price", null)]
        public string BurdenedUnitPrice { get; set; }

        [ColumnAttributes("Vendor Name", null)]
        public string VendorName { get; set; }

        [ColumnAttributes("Order Date", null)]
        public string OrderDate { get; set; }

        [ColumnAttributes("Required Delivery Date", null)]
        public string RequiredDeliveryDate { get; set; }

        [ColumnAttributes("Contract Number", null)]
        public string ContractNumber { get; set; }

        [ColumnAttributes("Order Number", null)]
        public string OrderNumber { get; set; }

        [ColumnAttributes("Award Number", null)]
        public string AwardNumber { get; set; }

        [ColumnAttributes("Comp Cost", null)]
        public string CompCost { get; set; }

        [ColumnAttributes("Extended Comp Cost", null)]
        public string ExtendedCompCost { get; set; }

        [ColumnAttributes("Minimum Sell Price", null, true)]
        public string MinimumSellPrice { get; set; }

        [ColumnAttributes("Deal Strategy", null, true)]
        public string DealStrategy { get; set; }

        [ColumnAttributes("ADS Cost", null, true)]
        public string ADSCost { get; set; }

        /// <summary>
        /// Parses the RFQ Quote Number to determine the fiscal year.
        /// </summary>
        /// <returns>String representing the fiscal year of the RFQ.</returns>
        private string DetermineRFQFiscalYear()
        {
            string quoteNumber = this.RFQQuoteNumber;
            string fiscalYear = (quoteNumber.Length >= 4 ?
                quoteNumber.Substring(0, 4) : // the rfq's fiscal year is denoted by the first 4 characters of the quote number
                "-"); // if there are less than 4 characters, then signal that the fiscal year is unknown
            return fiscalYear;
        }
    }
}
