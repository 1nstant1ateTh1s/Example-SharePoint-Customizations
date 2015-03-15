using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.SharePoint;

namespace RFQEventReceiver
{
    /// <summary>
    /// 
    /// </summary>
    public class AwardNotification : RFQOperation
    {
        // Properties
        public SPListItem RequestForQuoteListItem { get; set; }
        public string AwardNumber { get; set; }
        public double AwardAmount { get; set; }

        /// <summary>
        /// Represents an Award Notification document constructed from a newly added Award Notification List Item.
        /// </summary>
        /// <param name="awardNotificationListItem">The SharePoint List Item this object is representing.</param>
        public AwardNotification(SPListItem awardNotificationListItem) : base(awardNotificationListItem)
        {
            this.RFQQuoteNumber = ReadRFQNumber();
            this.RequestForQuoteListItem = FindRequestForQuoteListItem();
            this.AwardNumber = ReadAwardNumber();
            this.AwardAmount = ReadAwardAmount();
        }

        /// <summary>
        /// Handle processing of an Award Notification document.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        /// <param name="web">The current web context.</param>
        public void ProcessAwardNotificationData(SPItemEventProperties properties, SPWeb web)
        {
            try
            {
                RelocateAwardNotificationListItem(web); // move the award notification list item 
                UpdateAwardNotifListItem(); // update field values for the Award Notification library list item
                UpdateRequestForQuoteListItem(); // update field values for the related Request for Quote library list item
                UpdateRFQOrderItemsList(); // update rfq's order items
            }
            catch (Exception ex)
            {
                throw new RFQException(ex.Message, this.RFQQuoteNumber);
            }
        }

        /// <summary>
        /// Removes the Award Notification list item from the "Request for Quotes" library list & 
        /// adds it to the "Award Notifications" library list.
        /// </summary>
        /// <param name="web">The current web context.</param>
        private void RelocateAwardNotificationListItem(SPWeb web)
        {
            // Variables:
            SPListItem listItem = this.OperationListItem;

            try
            {
                // Move the award notification file from the Request for Quotes library list to
                // the Award Notifications library list:
                SPFile awardNotifFile = this.OperationFile;
                SPList destList = (SPDocumentLibrary)web.Lists.TryGetList(AWARD_NOTIFICATION_LIBRARY_NAME);
                string destUrl = destList.RootFolder.Url + "/" + listItem.File.Name;
                byte[] fileData = awardNotifFile.OpenBinary();
                
                // Add the Award Notification file to the designated document library, overwriting a file of the same name if one exists
                destList.RootFolder.Files.Add(destUrl, fileData, true);
            }
            catch (Exception ex)
            {
                throw new RFQException(ex.Message, this.RFQQuoteNumber);
            }
            finally
            {
                // Delete the award notif list item from the Request for Quotes library list
                listItem.Delete();

                // Find the new Award Notification library list item
                this.OperationListItem = FindAwardNotificationListItem(this.AwardNumber, web);
            }
        }

        /// <summary>
        /// Performs necessary updates to the Award Notification library list item.
        /// </summary>
        private void UpdateAwardNotifListItem()
        {
            SPListItem awardNotifListItem = this.OperationListItem;
            SPListItem requestForQuoteListItem = this.RequestForQuoteListItem;

            // Update field values
            awardNotifListItem["Award Number"] = this.AwardNumber;
            // Set Lookup Field value:
            SPFieldLookupValue rfqNum = new SPFieldLookupValue(requestForQuoteListItem.ID, this.RFQQuoteNumber);
            awardNotifListItem["RFQ"] = rfqNum;

            // Save changes to the award notification list item
            awardNotifListItem.SystemUpdate();
        }

        /// <summary>
        /// Performs necessary updates to the Request for Quotes library list item for fields that 
        /// require data resulting from parsing the Award Notification file's data.
        /// </summary>
        private void UpdateRequestForQuoteListItem()
        {
            SPListItem requestForQuoteListItem = this.RequestForQuoteListItem;
            SPListItem awardNotifListItem = this.OperationListItem;

            // Update field values
            requestForQuoteListItem["Status"] = RFQStatusTypes.AWARDED;
            requestForQuoteListItem["Award Amount"] = this.AwardAmount;
            // Set Lookup Field value:
            SPFieldLookupValue awardNum = new SPFieldLookupValue(awardNotifListItem.ID, this.AwardNumber);
            requestForQuoteListItem["Award Number"] = awardNum;

            // Save changes to the request for quote list item
            requestForQuoteListItem.SystemUpdate();
        }

        /// <summary>
        /// Update the RFQ Order Items based on Award Notification information.
        /// </summary>
        private void UpdateRFQOrderItemsList()
        {
            SPListItem requestForQuoteListItem = this.RequestForQuoteListItem;
            string listNm = requestForQuoteListItem["Contract Type"].ToString();
            SPList rfqOrderItemsList = requestForQuoteListItem.Web.Lists.TryGetList(listNm);
            string workflowName = "";

            // Update the RFQ Order Items w/ the Award Notification data
            PopulateRFQList(rfqOrderItemsList);

            switch (listNm)
            {
                case "SOE":
                    workflowName = "Declare SOE RFQ Order Item Record";
                    break;
                case "FES":
                    workflowName = "Declare FES RFQ Order Item Record";
                    break;
                case "TENT":
                    workflowName = "Declare TENT RFQ Order Item Record";
                    break;
                default:
                    break;
            }

            SPListItemCollection rfqOrderItems = FindRFQOrderItemsInList(rfqOrderItemsList);
            foreach (SPListItem item in rfqOrderItems)
            {
                // Call workflow to move list item to archive list
                Common.StartWorkflow(item, workflowName);
            }
        }

        /// <summary>
        /// Parse's the Award Notification file into the appropriate custom SharePoint list.
        /// </summary>
        /// <param name="rfqList">The target SharePoint list with which to populate the Award Notification file's data.</param>
        private void PopulateRFQList(SPList rfqList)
        {
            // Variables
            SPListItem awardNotifListItem = this.OperationListItem;
            DataTable awardNotifData = this.OperationFileData;
            SPFieldLookupValue awardNum = new SPFieldLookupValue(awardNotifListItem.ID, this.AwardNumber);
            string orderItemId;

            foreach (DataRow dr in awardNotifData.Rows)
            {
                // Check to make sure the list already contains the order item for this RFQ
                orderItemId = dr["Order Item ID"].ToString();
                SPListItem orderListItem = FindOrderItemInList(rfqList, orderItemId);
                if (orderListItem != null)
                {
                    // Set list item field values:
                    orderListItem["Line Status"] = OrderItemStatusTypes.AWARDED; // set "Won" status for ea. of the affected order items

                    orderListItem["Purchase Unit Price"] = dr["Purchase Unit Price"];
                    orderListItem["Purchase Extended Price"] = dr["Purchase Extended Price"];
                    orderListItem["Burdened Unit Price"] = dr["Burdened Unit Price"];

                    orderListItem["Award Number"] = awardNum; // set LookupField value

                    orderListItem["Order Date"] = ExcelDocumentUtil.ReadExcelDateTimeValue(dr["Order Date"]);
                    orderListItem["Required Delivery Date"] = ExcelDocumentUtil.ReadExcelDateTimeValue(dr["Required Delivery Date"]);
                    orderListItem["Contract Number"] = dr["Contract Number"];
                    orderListItem["Order Number"] = dr["Order Number"];

                    orderListItem.Update(); // save changes to list item
                }

            }
        }

        /// <summary>
        /// Retrieve the RFQ Quote Number from within the Award Notification file data.
        /// </summary>
        /// <returns>The Request for Quote ID number.</returns>
        private string ReadRFQNumber()
        {
            string rfqNum = "";
            DataTable dt = this.OperationFileData;
            
            // Check to see if there is any data to read
            if (dt.Rows.Count > 0)
            {
                // Find the RFQ's quote number (Note: all rows have the same rfq quote number)
                DataRow dr = dt.Rows[0];
                try
                {
                    rfqNum = dr["RFQ Number"].ToString();
                }
                catch (Exception)
                {
                    rfqNum = "Unknown";
                }
            }

            return rfqNum; // return the rfq quote #
        }

        /// <summary>
        /// Retrieve the Award Number from within the Award Notification file data.
        /// </summary>
        /// <returns>The Award Notification's award number.</returns>
        private string ReadAwardNumber()
        {
            string awardNum = "";
            DataTable dt = this.OperationFileData;

            // Check to see if there is any data to read
            if (dt.Rows.Count > 0)
            {
                // Find the Award Notification's award number (Note: all rows refer to the same award number)
                DataRow dr = dt.Rows[0];
                awardNum = dr["Award Number"].ToString();
            }

            return awardNum; // return the award #
        }

        /// <summary>
        /// Retrieve the Award Amount total from within the Award Notification file data.
        /// </summary>
        /// <returns>The Award Amount total.</returns>
        private double ReadAwardAmount()
        {
            double awardAmount = 0.00;
            DataTable dt = this.OperationFileData;

            // Loop through ea. row of data and total the value for the 'Purchase Extended Price' column
            foreach (DataRow dr in dt.Rows)
            {
                awardAmount += double.Parse(dr["Purchase Extended Price"].ToString());
            }

            return awardAmount;
        }

    }
}
