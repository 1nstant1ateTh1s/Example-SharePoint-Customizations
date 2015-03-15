using System;
using System.IO.Packaging;
using System.IO;
using System.Linq;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Collections.Specialized;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RFQEventReceiver.RFQEventReceiver
{
    delegate void ProcessRFQList(SPItemEventProperties properties, SPWeb web);

    /// <summary>
    /// List Item Events
    /// </summary>
    public class RFQEventReceiver : SPItemEventReceiver
    {
        /// <summary>
        /// An item is being added
        /// </summary>
        public override void ItemAdding(SPItemEventProperties properties)
        {
            base.ItemAdding(properties);

            // Prevent invalid file types from being uploaded to the document library
            checkValidFileType(properties);
        }

        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);

            // Parse RFQ Excel document into custom SharePoint list unless the RFQ Order Items already exist.
            extractRFQDataEvent(properties);
        }

        /// <summary>
        /// An item is being updated
        /// </summary>
        public override void ItemUpdating(SPItemEventProperties properties)
        {
            base.ItemUpdating(properties);
        }

        /// <summary>
        /// An item was updated
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);

            // If Request for Quote Item status is updated to either "Awarded" or "Loss", close all of respective active RFQ Order Items.
            string beforeStatus = (properties.BeforeProperties["Status"] != null ? properties.BeforeProperties["Status"].ToString() : "");
            string afterStatus = (properties.AfterProperties["Status"] != null ? properties.AfterProperties["Status"].ToString() : "");
                        
            if ((afterStatus == RFQStatusTypes.LOSS && beforeStatus != RFQStatusTypes.LOSS) || // compare before & after status values so that the attempt to close Order Items only occurs
                (afterStatus == RFQStatusTypes.AWARDED && beforeStatus != RFQStatusTypes.AWARDED)) // on the initial "Status" set to "Awarded" or "Loss" & not on subsequent updates w/ no changes to "Status" occur
            {
                // When a Request for Quote status is manually updated to "Awarded" or "Loss", close out that RFQ's respective active Order Items.
                closeRFQOrderItemsEvent(properties);
            }
        }

        /// <summary>
        /// Determine validity of uploaded document. Only Excel files in the OpenXML-compatible .xlsx format are allowed.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        private void checkValidFileType(SPItemEventProperties properties)
        {
            // Check file type of uploaded document - only Excel files in the OpenXML-compatible .xlsx format are to be stored in the list
            if (!(properties.AfterUrl).ToLower().Contains(".xlsx"))
            {
                properties.Status = SPEventReceiverStatus.CancelNoError;
                properties.Cancel = true; // cancel upload
            }
        }

        /// <summary>
        /// Handle the custom actions necessary for a Request For Quote ItemAdded event.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        private void extractRFQDataEvent(SPItemEventProperties properties)
        {
            // Set site context
            using (SPSite site = new SPSite(properties.WebUrl))
            {
                // Set web context
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        SPListItem listItem = properties.ListItem;
                        string listItemNm = listItem.Name;
                        ProcessRFQList processRFQList = null;

                        // Determine appropriate action to take depending upon the file name of the uploaded document. Store the action to take in delegate.
                        if (listItemNm.Contains(FileType.REQUEST_FOR_QUOTE))
                        {
                            // Handle an incoming Request for Quote document
                            RequestForQuote rfq = new RequestForQuote(listItem);
                            processRFQList = rfq.ProcessRequestForQuote;
                        }
                        else if (listItemNm.Contains(FileType.AWARD_NOTIFICATION))
                        {
                            // Handle an incoming Award Notification document
                            AwardNotification awardNotif = new AwardNotification(listItem);
                            processRFQList = awardNotif.ProcessAwardNotificationData;
                        }

                        // Parse new RFQ Excel document into custom SharePoint list
                        if (processRFQList != null)
                        {
                            EventFiringEnabled = false;
                            try
                            {
                                processRFQList(properties, web); // execute delegate
                            }
                            catch (RFQException rfqEx)
                            {
                                properties.ErrorMessage = rfqEx.Message;
                                logItemEventProperties(properties, rfqEx.RFQNumber);
                            }
                            finally
                            {
                                EventFiringEnabled = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        properties.ErrorMessage = ex.Message;
                        logItemEventProperties(properties, "");
                    }

                }
            }
        }

        /// <summary>
        /// Closes out a Request for Quote's respective active Order Items based on the Request for Quote's awarded or loss status.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        private void closeRFQOrderItemsEvent(SPItemEventProperties properties)
        {
            // Set site context
            using (SPSite site = new SPSite(properties.WebUrl))
            {
                // Set web context
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        // Close RFQ Order Items:
                        RequestForQuote reqForQuote = new RequestForQuote(properties);
                        reqForQuote.CloseRFQOrderItems();

                    }
                    catch (SPException ex)
                    {
                        string rfqQuoteNum = properties.ListItem["RFQ Quote Number"].ToString();
                        properties.ErrorMessage = ex.Message;
                        logItemEventProperties(properties, rfqQuoteNum);
                    }
                }
            }

        }

        /// <summary>
        /// Log list item event properties
        /// </summary>
        /// <param name="properties">The list item event properties from which to gather the details to be logged.</param>
        /// <param name="rfqQuoteNum"></param>
        private void logItemEventProperties(SPItemEventProperties properties, string rfqQuoteNum)
        {
            // Specify the Logs list name
            string logName = "RFQ Event Logger";

            // Create the stringbuilder object
            StringBuilder sb = new StringBuilder();

            // Add properties that do not throw an exception
            sb.AppendFormat("ErrorMessage: {0}\n", properties.ErrorMessage);
            
            // Log the event to the list
            this.EventFiringEnabled = false;
            Common.LogEvent(properties.Web, logName, properties.EventType, rfqQuoteNum, sb.ToString());
            this.EventFiringEnabled = true;
        }

    }
}
