using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceModel.Activation;
using System.Data;
using System.Reflection;
using System.Net.Mail;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client.Services;
using Microsoft.SharePoint.Administration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RFQ_SharePoint_Project.WebServices
{
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    public class RFQWorkflowService : IRFQWorkflowService
    {

        public JsonResponse GenRFQReturnFileSOE(SOE_RFQOrderItems orderItems, string relativeUrl)
        {
            return GenRFQReturnFile(orderItems, relativeUrl);
        }

        public JsonResponse GenRFQReturnFileFES(FES_RFQOrderItems orderItems, string relativeUrl)
        {
            return GenRFQReturnFile(orderItems, relativeUrl);
        }

        public JsonResponse GenRFQReturnFileTENT(TENT_RFQOrderItems orderItems, string relativeUrl)
        {
            return GenRFQReturnFile(orderItems, relativeUrl);
        }

        /// <summary>
        /// Parses a series of RFQ Order Items & exports the data to the relevant RFQ Excel file in the "Request for Quotes"
        /// SharePoint Document Library. This generates the RFQ Workflow return file.
        /// </summary>
        /// <param name="orderItems">The series of RFQ Order Item's data.</param>
        /// <param name="relativeUrl">The current relative url for the active site.</param>
        /// <returns>An indicator of whether the operation was a success.</returns>
        public JsonResponse GenRFQReturnFile<T>(RFQOrderItems<T> orderItems, string relativeUrl) where T : RFQOrderItem
        {

            // Variables
            JsonResponse returnObj = new JsonResponse() { IsSuccess = false };
            string serverUrl = SPContext.Current.Site.Url;
            string siteUrl = serverUrl + relativeUrl;
            string rfqQuoteNum = orderItems[0].Quote_Number.ToString();
            string libraryName = "Request for Quotes";

            try
            {
                // Set site context
                using (SPSite site = new SPSite(siteUrl))
                {
                    // Set web context
                    using (SPWeb web = site.OpenWeb())
                    {
                        // Search the document library for the RFQ Excel file that matches the current RFQ Quote Number being operated on
                        SPFile rfqFile = findRFQExcelFile(web, libraryName, rfqQuoteNum);
                        
                        if (rfqFile != null) // check that a file object was returned
                        {
                            // Temporarily allow unsafe updates
                            web.AllowUnsafeUpdates = true;

                            // Write data to the RFQ Excel file
                            exportRFQDataToExcelFile(rfqFile, orderItems);
                            returnObj.IsSuccess = true;

                            if (returnObj.IsSuccess) // if data was written successfully, update 'Request for Quotes' Library list item fields:
                            {
                                // Status is set to a variance of the "Approval" status based upon the RFQ's total quote dollar value
                                SPListItem rfqListItem = GetRFQListItem(web, libraryName, rfqQuoteNum);
                                double totalQuoteVal = CalcTotalQuoteValue(orderItems);

                                rfqListItem["Status"] = RFQStatusTypes.GetApprovalStatusType(totalQuoteVal);
                                rfqListItem["ADS Bid"] = totalQuoteVal.ToString();
                                rfqListItem.Update(); // save changes to list item
                                
                            }

                            // Revert AllowUnsafeUpdates back to false
                            web.AllowUnsafeUpdates = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Record error message to be returned to client
                returnObj.Errors.Add(ex.Message.ToString());
            }

            return returnObj; // return response object to client 
        }

        /// <summary>
        /// Submits the RFQ Bid that corresponds to the specified List Item ID via e-mail to the specified inbox account.
        /// </summary>
        /// <param name="rfq">The ID of the RFQ list item to submit.</param>
        /// <param name="relativeUrl">The current relative url for the active site.</param>
        /// <returns>An indicator of whether the operation was a success.</returns>
        public JsonResponse SubmitRFQBid(int rfqItemId, string relativeUrl)
        {
            JsonResponse returnObj = new JsonResponse() { IsSuccess = false };
            string serverUrl = SPContext.Current.Site.Url;
            string siteUrl = serverUrl + relativeUrl;
            string libraryName = "Request for Quotes";

            try
            {
                // Set site context
                using (SPSite site = new SPSite(siteUrl))
                {
                    // Set web context
                    using (SPWeb web = site.OpenWeb())
                    {
                        // Attempt to retrieve the ListItem by ID
                        SPList list = web.Lists.TryGetList(libraryName);
                        SPListItem rfqListItem = list.GetItemById(rfqItemId);

                        if (rfqListItem != null)
                        {
                            // If the list item was found, send e-mail containing the relevant RFQ list item data
                            sendEmailWithAttachment(rfqListItem, web);
                            returnObj.IsSuccess = true; // if no exceptions encountered, set success property to true

                            if (returnObj.IsSuccess) // if bid was successfully emailed, update status for 'Request for Quote' list item
                            {
                                web.AllowUnsafeUpdates = true; // temporarily allow unsafe updates                                
                                updateListItemStatus(rfqListItem, RFQStatusTypes.SUBMITTED);
                                web.AllowUnsafeUpdates = false; // revert AllowUnsafeUpdates back to false
                            }
                        }
                    }
                }
            }
            catch (SPException ex)
            {
                // Record error message to be returned to client
                returnObj.IsSuccess = false;
                returnObj.Errors.Add(ex.Message.ToString());
            }

            return returnObj; // return response object to client
        }

        /// <summary>
        /// Retrieves the Request for Quote list item that matches the specified rfq quote number & list name.
        /// </summary>
        /// <param name="web">The active SharePoint Web context.</param>
        /// <param name="listName">The name of the list in which to search.</param>
        /// <param name="rfqQuoteNum">The RFQ Quote Number.</param>
        /// <returns>If found, return a list item that matches the specified rfq quote number; otherwise returns null.</returns>
        private SPListItem GetRFQListItem(SPWeb web, string listName, string rfqQuoteNum)
        {
            // Variables: 
            SPList rfqLib = web.Lists.TryGetList(listName);
            SPListItemCollection rfqLibItemCollection;
            SPListItem rfqLibItem = null;
            SPQuery queryObj = new SPQuery();
            string queryString = @"<Where><Contains><FieldRef Name='BaseName'/>" +
                "<Value Type='Text'>" + rfqQuoteNum + "</Value>" +
                "</Contains></Where>";

            // Execute query
            queryObj.Query = queryString;
            rfqLibItemCollection = rfqLib.GetItems(queryObj);

            // Check if list item was found
            if (rfqLibItemCollection.Count > 0) // list item found for the specified RFQ Quote Number
            {
                // Grab list item from 1st position
                rfqLibItem = rfqLibItemCollection[0];
            }

            return rfqLibItem; // return list item
        }

        /// <summary>
        /// Totals the Purchase Extended Price value for ea. of the specified order items to determine the total quote price.
        /// </summary>
        /// <param name="orderItems">The collection of RFQ Order Items.</param>
        /// <returns>The total quote price for the collection of order items.</returns>
        private double CalcTotalQuoteValue<T>(RFQOrderItems<T> orderItems)
        {
            // Variables
            double result = 0;
            int i = 0, count = orderItems.Count;
            PropertyInfo propInfo = typeof(T).GetProperty("Purchase_Extended_Price");

            // Loop through & total ea. order item's 'Purchase Extended Price' value
            for (; i < count; i++ )
            {
                result += double.Parse(propInfo.GetValue(orderItems[i], null).ToString());
            }
            return result;
        }

        /// <summary>
        /// Sends an E-mail with data from the specified RFQ List Item.
        /// </summary>
        /// <param name="listItem">The RFQ List Item to be E-mailed.</param>
        /// <param name="web">The active SharePoint Web context.</param>
        private void sendEmailWithAttachment(SPListItem listItem, SPWeb web)
        {
            try
            {
                List<string> recipients = new List<string>();
                recipients.Add("microlink@adsinc.com");

                SPFile file = listItem.File; // the file will be added as an attachment to the e-mail

                // Create the MailMessage & supply it w/ the relevant information
                MailMessage message = new MailMessage();
                message.From = new MailAddress(SPAdministrationWebApplication.Local.OutboundMailSenderAddress);
                foreach (var recip in recipients)
                {
                    message.To.Add(recip.ToString());
                }
                message.Subject = "RFQ: # - Bid Submission";
                message.Body = "Test RFQ Bid Submission.";
                message.Attachments.Add(new Attachment(file.OpenBinaryStream(), file.Name)); // add the attachment

                // Create the SMTP client object & send the message
                // SmtpClient class sends the email by using the specified SMTP server
                SmtpClient smtpClient = new SmtpClient(SPAdministrationWebApplication.Local.OutboundMailServiceInstance.Server.Address);
                smtpClient.Send(message);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Given a SharePoint List Item, attempts to update that list item's status field to the provided status value.
        /// </summary>
        /// <param name="rfqLibItem"></param>
        /// <param name="status">Represents the value to update the list item's status field to.</param>
        private void updateListItemStatus(SPListItem rfqLibItem, string status)
        {
            // Update field value
            rfqLibItem["Status"] = status;

            // Save changes to the current list item
            rfqLibItem.Update();
        }
        
        /// <summary>
        /// Retrieves the RFQ Excel file from the specified SharePoint web site & library/list based on the provided RFQ Quote Number.
        /// </summary>
        /// <param name="web">The SharePoint Web object in which to search for the specified library.</param>
        /// <param name="libraryName">The name of the document library/list containing the RFQ Excel files.</param>
        /// <param name="rfqQuoteNum">The active RFQ Quote Number. Used to search the filenames to find the relevant file.</param>
        /// <returns>The relevant File object representing the stored RFQ Excel file.</returns>
        private SPFile findRFQExcelFile(SPWeb web, string libraryName, string rfqQuoteNum)
        {
            // Variables:
            SPList rfqLib = web.Lists.TryGetList(libraryName);
            SPListItemCollection rfqFileCollection;
            SPListItem rfqFileItem;
            SPFile rfqFile;
            SPQuery queryObj = new SPQuery();
            string queryString = @"<Where><Contains><FieldRef Name='BaseName'/>" +
                "<Value Type='Text'>" + rfqQuoteNum + "</Value>" +
                "</Contains></Where>";

            try
            {
                // Execute query
                queryObj.Query = queryString;
                rfqFileCollection = rfqLib.GetItems(queryObj);

                // Check if item was found
                if (rfqFileCollection.Count > 0) // item found for the specified RFQ Quote Number
                {
                    // Grab file from 1st position
                    rfqFileItem = rfqFileCollection[0];
                    rfqFile = rfqFileItem.File;
                }
                else // no item found for the specified RFQ Quote Number
                {
                    throw new SPException("No file found in SharePoint Document Library: " + libraryName + " for RFQ #" + rfqQuoteNum + ".");
                }

                // Return file object
                return rfqFile;
            }
            catch (SPException ex)
            {
                throw ex;
            }
        }
        
        /// <summary>
        /// Given a target Excel SharePoint file object & a collection of RFQ Order Items data, populates a row in the 
        /// Excel file for each of the provided data items.
        /// </summary>
        /// <param name="rfqFile">The Excel file to write to.</param>
        /// <param name="orderItems">The collection of data used to populate the file.</param>
        private void exportRFQDataToExcelFile<T>(SPFile rfqFile, RFQOrderItems<T> orderItems)
        {
            // Variables
            Stream dataStream;
            try
            {                
                // Create binary stream for opening the file
                dataStream = rfqFile.OpenBinaryStream();

                // Open the spreadsheet document w/ write access
                using (SpreadsheetDocument doc =
                    SpreadsheetDocument.Open(dataStream, true))
                {
                    // Retrieve references
                    WorkbookPart wbPart = doc.WorkbookPart;
                    WorksheetPart wsPart = wbPart.WorksheetParts.LastOrDefault(); // multiple worksheets are in decending order - grabbing the last position will retrieve the "first" sheet
                    SharedStringTablePart ssTblPart = wbPart.SharedStringTablePart;

                    wbPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    wbPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                    // Ensure worksheet part exists
                    if (wsPart != null)
                    {
                        SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                        int orderItemCount = orderItems.Count;
                        PropertyInfo[] props = typeof(T).GetProperties();
                        List<T> items = orderItems.ToList<T>();

                        // Use Reflection to get property information for the RFQ Order Items
                        foreach (PropertyInfo info in props)
                        {
                            string propertyName = info.Name.Replace("_", " "); // remove any underscore characters from property names in order to facilitate matching to excel column names
                            string columnName = ExcelDocumentUtil.GetColumnReference(propertyName, wsPart, ssTblPart); // given a property name, looks up the relevant column index (i.e., 'A', 'B', 'Z', 'AA', ...)

                            if (columnName != null)
                            {
                                // For each property/column, we are going to iterate over ea. RFQOrderItem object, 
                                // so cells will be populated like so: 'A2', 'A3', 'An'; 'B2', 'B3', 'Bn'; 'C2', 'C3', ...
                                for (var i = 0; i < orderItemCount; i++)
                                {
                                    uint rowIndex = ((uint)(i + 2)); // Row 1 is used for column headings, so start row index at position 2
                                    Cell targetCell = ExcelDocumentUtil.InsertCellInWorksheet(columnName, rowIndex, wsPart); // inserts a cell at the given position, or just returns it if one already exists

                                    var cellValue = info.GetValue(items[i], null); // gets the value of this iteration's property from the current RFQOrderItem object
                                    ExcelDocumentUtil.AddValueToCell(targetCell, cellValue, ssTblPart); // add the value to the cell
                                }
                            }
                            else
                            {

                            }
                        }

                        // Once all the data is written into the excel sheet, 
                        // flush any cells that contain a formula to force recalculation of their value based on the new data ...                        
                        ExcelDocumentUtil.FlushCachedValues(doc);                        
                        // ... & save the changes
                        wsPart.Worksheet.Save();
                        rfqFile.SaveBinary(dataStream); // save file
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

    }
}
