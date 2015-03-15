using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.SharePoint;

namespace RFQEventReceiver
{
    public class RFQOperation
    {
        // Constants
        public const string RFQ_LIBRARY_NAME = "Request for Quotes";
        public const string AWARD_NOTIFICATION_LIBRARY_NAME = "Award Notifications";
        public const string RFQ_HISTORY_LIST = "QUOTATION_HISTORY";
        public const string RFQ_EVENT_LOGGER_LIST = "RFQ Event Logger";

        // Properties
        public string RFQQuoteNumber { get; set; }
        protected SPListItem OperationListItem { get; set; }
        protected SPFile OperationFile { get; set; }
        protected DataTable OperationFileData { get; set; }
        
        private int counter = 0; // for testing purposes

        /// <summary>
        /// Default constructor
        /// </summary>
        public RFQOperation()
        {
        }

        /// <summary>
        /// Represents a base general RFQ Workflow operation from which specific operations are derived constructed from 
        /// a newly added Request for Quote List Item.
        /// </summary>
        /// <param name="newListItem">The SharePoint List Item that will be represented as an operation.</param>
        public RFQOperation(SPListItem newListItem)
        {
            this.OperationListItem = newListItem;
            this.OperationFile = newListItem.File;
            this.OperationFileData = ExcelDocumentUtil.ExtractExcelSheetToDataTable(newListItem.File);
        }
        
        /// <summary>
        /// Searches within a SharePoint List to determine if an item already exists or not. 
        /// An item is uniquely identified by the Order Item ID & the Quote Number.
        /// </summary>
        /// <param name="list">The SharePoint List on which to search.</param>
        /// <param name="orderItemId">The order item id.</param>
        /// <param name="quoteNumber">The RFQ quote number.</param>
        /// <returns>True if list item already exists, false if it does not.</returns>
        protected bool OrderItemAlreadyExists(SPList list, string orderItemId, string quoteNumber)
        {
            // Variables
            SPListItemCollection listItem = null;
            SPQuery queryObj = new SPQuery();
            string stringQuery = @"<Where><And><Eq><FieldRef Name='Title'/><Value Type='Text'>" + orderItemId + "</Value>" +
                "</Eq><Eq><FieldRef Name='Quote_Number'/><Value Type='Text'>" + quoteNumber + "</Value></Eq></And></Where>";

            if (list != null)
            {
                // Execute query
                queryObj.Query = stringQuery;
                listItem = list.GetItems(queryObj);
            }

            // Return true if item already exists in list item collection, otherwise false
            return (listItem.Count > 0);
        }

        /// <summary>
        /// Attempts to retrieve the Request for Quotes Document Library list item that matches the current RFQ Quote Number.
        /// </summary>
        /// <returns>The Request for Quotes Document Library list item with the specified RFQ Quote Number.</returns>
        protected SPListItem FindRequestForQuoteListItem()
        {
            // Variables
            SPList requestForQuoteList = this.OperationListItem.Web.Lists.TryGetList(RFQ_LIBRARY_NAME);
            SPListItem requestForQuoteListItem = null;
            SPListItemCollection listItems;
            SPQuery queryObj = new SPQuery();
            string stringQuery = @"<Where><Eq>" +
                "<FieldRef Name='RFQ_Quote_Number'/><Value Type='Text'>" + this.RFQQuoteNumber + "</Value>" +
                "</Eq></Where>";

            if (requestForQuoteList != null)
            {
                // Execute query
                queryObj.Query = stringQuery;
                listItems = requestForQuoteList.GetItems(queryObj);

                // Check if item was found
                if (listItems.Count > 0) // item found for the specified RFQ Quote Number
                {
                    requestForQuoteListItem = listItems[0]; // grab list item from 1st position
                }
            }
            else
            {
                string logMsg = RFQ_LIBRARY_NAME + " list not found. " + "Hist list lookup counter: " + counter++;
                Common.LogEvent(this.OperationListItem.Web, RFQ_EVENT_LOGGER_LIST, this.RFQQuoteNumber, logMsg);
            }

            return requestForQuoteListItem; // return list item
        }

        /// <summary>
        /// Attempts to retrieve the Award Notification Document Library list item that matches the specified Award Number.
        /// </summary>
        /// <param name="awardNumber">The Award Number to search for.</param>
        /// <param name="web">The current web context.</param>
        /// <returns>The Award Notification Document Library list item w/ the specified Award Number.</returns>
        protected SPListItem FindAwardNotificationListItem(string awardNumber, SPWeb web)
        {
            // Variables
            SPList awardNotificationList = web.Lists.TryGetList(AWARD_NOTIFICATION_LIBRARY_NAME);
            SPListItemCollection listItems;
            SPListItem awardNotifLibItem = null;
            SPQuery queryObj = new SPQuery();
            string queryString = @"<Where><Contains>" +
                "<FieldRef Name='BaseName'/><Value Type='Text'>" + awardNumber + "</Value>" +
                "</Contains></Where>";

            if (awardNotificationList != null)
            {
                // Execute query:
                queryObj.Query = queryString;
                listItems = awardNotificationList.GetItems(queryObj);

                // Check if item was found
                if (listItems.Count > 0) // list item found for the specified Award Number
                {
                    awardNotifLibItem = listItems[0]; // grab list item from 1st position
                }
            }
            else
            {
                string logMsg = AWARD_NOTIFICATION_LIBRARY_NAME + " list not found.";
                Common.LogEvent(web, RFQ_EVENT_LOGGER_LIST, this.RFQQuoteNumber, logMsg);
            }

            return awardNotifLibItem; // return list item
        }

        /// <summary>
        /// Attempts to retrieve a RFQ History list item that matches the current RFQ Manufacturer Part Number.
        /// </summary>
        /// <param name="manufacturerPartNm">The Manufacturer Part Number to search for.</param>
        /// <returns>The RFQ History list item if one exists.</returns>
        protected SPListItem FindRFQHistoryListItem(string manufacturerPartNm, string quantity)
        {
            // Variables
            SPList rfqHistoryList = this.OperationListItem.Web.Lists.TryGetList(RFQ_HISTORY_LIST);
            SPListItemCollection listItems;
            SPListItem rfqHistoryListItem = null;            
            SPQuery queryObj = new SPQuery();
            string queryString = @"<Where><And><Eq>" +
                "<FieldRef Name='Title'/><Value Type='Text'>" + manufacturerPartNm + "</Value>" +
                "</Eq><Eq>" + 
                "<FieldRef Name='Quantity'/><Value Type='Number'>" + quantity + "</Value>" +
                "</Eq></And></Where>";

            if (rfqHistoryList != null) // if a list was found
            {
                // Execute query
                queryObj.Query = queryString;
                listItems = rfqHistoryList.GetItems(queryObj);

                // Check if item was found
                if (listItems.Count > 0) // item found for the specified RFQ Manufacturer Part Number
                {
                    rfqHistoryListItem = listItems[0]; // grab list item from 1st position
                }
            }
            else
            {
                string logMsg = RFQ_HISTORY_LIST + " list not found. " + "Hist list lookup counter: " + counter++;
                Common.LogEvent(this.OperationListItem.Web, RFQ_EVENT_LOGGER_LIST, this.RFQQuoteNumber, logMsg);
            }

            return rfqHistoryListItem; // return list item
        }

        /// <summary>
        /// Given a SharePoint list, attempts to retrieve a collection of list items that match the current Quote Number.
        /// </summary>
        /// <param name="list">The SharePoint List on which to search.</param>
        /// <returns>A collection of list items that match the current RFQ Quote Number within the specified SharePoint List.</returns>
        protected SPListItemCollection FindRFQOrderItemsInList(SPList list)
        {
            // Variables
            SPListItemCollection orderItems = null;
            SPQuery queryObj = new SPQuery();
            string stringQuery = @"<Where><Eq>" +
                "<FieldRef Name='Quote_Number'/><Value Type='Text'>" + this.RFQQuoteNumber + "</Value>" +
                "</Eq></Where>";

            if (list != null)
            {
                // Execute query
                queryObj.Query = stringQuery;
                orderItems = list.GetItems(queryObj);
            }

            return orderItems; // return list item collection
        }

        /// <summary>
        /// Given a SharePoint list, attempts to retrieve a list item that matches the given Order Item ID.
        /// </summary>
        /// <param name="list">The SharePoint List on which to search.</param>
        /// <param name="orderItemId">The order item id.</param>
        /// <returns>A list item with the specified Order Item ID value.</returns>
        protected SPListItem FindOrderItemInList(SPList list, string orderItemId)
        {
            // Variables
            SPListItemCollection orderItems = null;
            SPQuery queryObj = new SPQuery();
            string stringQuery = @"<Where><Eq>" +
                "<FieldRef Name='Title'/><Value Type='Text'>" + orderItemId + "</Value>" +
                "</Eq></Where>";

            if (list != null)
            {
                // Execute query
                queryObj.Query = stringQuery;
                orderItems = list.GetItems(queryObj);
            }

            return (orderItems.Count > 0 ? orderItems[0] : null); // return list item if found, otherwise null
        }

    }
}
