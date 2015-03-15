using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using Microsoft.SharePoint;
using RFQEventReceiver.Entities.OrderItems;

namespace RFQEventReceiver
{
    /// <summary>
    /// 
    /// </summary>
    public class RequestForQuote : RFQOperation
    {
        // Properties
        public string ContractNm { get; set; }
        public string Status { get; set; }
        public DateTime QuoteDueDate { get; set; }
        public string SalesRep { get; set; }
        public SPList RFQOrderItemsActiveList { get; set; }
        public SPList RFQOrderItemsArchiveList { get; set; }
        
        public List<OrderItemBase> OrderItems { get; set; }


        /// <summary>
        /// Represents a Request for Quote constructed from a list item's event properties that already exist.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        public RequestForQuote(SPItemEventProperties properties)
        {
            this.RFQQuoteNumber = properties.ListItem["RFQ Quote Number"].ToString();
            this.ContractNm = properties.ListItem["Contract Type"].ToString();
            this.Status = properties.ListItem["Status"].ToString();
            this.RFQOrderItemsActiveList = GetTargetOrderItemsActiveList(properties.ListItem.Web);
            this.RFQOrderItemsArchiveList = GetTargetOrderItemsArchiveList(properties.ListItem.Web);
        }

        /// <summary>
        /// Represents a Request for Quote document constructed from a newly added Request for Quote List Item.
        /// </summary>
        /// <param name="requestForQuoteListItem">The SharePoint List Item this object is representing.</param>
        public RequestForQuote(SPListItem requestForQuoteListItem) : base(requestForQuoteListItem)
        {
            this.ContractNm = DetermineRFQContractType();
            this.OrderItems = ReadRFQOrderItems();

            this.RFQQuoteNumber = this.OrderItems[0].RFQQuoteNumber;
            this.QuoteDueDate = ExcelDocumentUtil.ReadExcelDateTimeValue(this.OrderItems[0].QuoteDueDate); // set DateTime value

            //this.SalesRep = LookupSalesRep();    // NOT CURRENTLY WORKING
            this.Status = DetermineRFQStatus();
            this.RFQOrderItemsActiveList = GetTargetOrderItemsActiveList(requestForQuoteListItem.Web);
            this.RFQOrderItemsArchiveList = GetTargetOrderItemsArchiveList(requestForQuoteListItem.Web); 
        }

        /// <summary>
        /// Reads the Request for Quote Order Items.
        /// </summary>
        /// <returns>The list of Order Items for the current RFQ.</returns>
        private List<OrderItemBase> ReadRFQOrderItems()
        {
            List<OrderItemBase> orderItems = new List<OrderItemBase>();
            DataTable dt = this.OperationFileData;
            OrderItemBase orderItem = OrderItemFactory.CreateOrderItem(this.ContractNm); // simple factory pattern to retrieve correct type of Order Item class
            int count = 0; // track the row number
            int totalCount = dt.Rows.Count; // track total number of items within a series on each item

            if (orderItem != null)
            {
                object[,] oProperties = orderItem.GetProperties(orderItem.GetType());

                foreach (DataRow dr in dt.Rows) // read ea. order item row
                {
                    orderItem = OrderItemFactory.CreateOrderItem(this.ContractNm); // create new order item object
                    orderItem.SetExcelProperties(oProperties, dr); // populate order item's properties

                    // Assign an index number to represent the item's position within the series of items
                    orderItem.SeriesNumber = ++count; // increment counter and assign to list item
                    orderItem.SeriesTotal = totalCount; // track total number of items within this series

                    // Lookup any relevant historical data 
                    LookupOrderItemHistoryData(orderItem);

                    orderItems.Add(orderItem); // add to list of order items
                }
            }

            return orderItems; // return list of order items
        }

        /// <summary>
        /// Handle processing of a Request for Quote document.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        /// <param name="web">The current web context.</param>
        public void ProcessRequestForQuote(SPItemEventProperties properties, SPWeb web)
        {
            try
            {
                UpdateListItem(); // set field values for the Request for Quote library list item
                PopulateRFQOrderItemsList(); // populate the rfq order item list data
            }
            catch (Exception ex)
            {             
                throw new RFQException(ex.Message, this.RFQQuoteNumber);
            }
        }

        /// <summary>
        /// Closes out a Request for Quote's respective active Order Items based on the Request for Quote's "awarded" or "loss" status.
        /// </summary>
        /// <param name="properties">Properties for the list item event.</param>
        /// <param name="web">The current web context.</param>
        public void CloseRFQOrderItems()
        {
            try
            {                
                // Set Order Item status based on current Request for Quote item status:                
                SPListItemCollection rfqOrderItems = FindRFQOrderItemsInList(this.RFQOrderItemsActiveList); // find active order items that match the given RFQ
                string orderItemLineStatus = OrderItemStatusTypes.TranslateStatus(this.Status);

                foreach (SPListItem item in rfqOrderItems)
                {
                    // Update list item line status value:
                    item["Line Status"] = orderItemLineStatus;
                    item.Update(); // save changes to list
                }

            }
            catch (SPException ex)
            {
                throw new RFQException(ex.Message, this.RFQQuoteNumber);
            }
        }

        /// <summary>
        /// Performs necessary updates to the Request for Quotes library list item for fields that 
        /// require data resulting from parsing the RFQ file's data.
        /// </summary>
        private void UpdateListItem()
        {
            SPListItem requestForQuoteListItem = this.OperationListItem;
             
            // Set field values
            requestForQuoteListItem["RFQ Quote Number"] = this.RFQQuoteNumber;
            requestForQuoteListItem["Contract Type"] = this.ContractNm;
            requestForQuoteListItem["Status"] = this.Status;
            requestForQuoteListItem["Quote Due Date"] = this.QuoteDueDate;

            // Handle special case of setting Person/Group field value:
            /* *** NOTE: TECHNICAL LIMITATION HERE. MUST ASSIGN A SPFieldUserValue TO A People/Groups FIELD. *** */
            //SPFieldUserValue salesRepUser = new SPFieldUserValue(requestForQuoteListItem.Web, this.SalesRep);
            //requestForQuoteListItem["Sales Rep"] = salesRepUser;            

            // Save changes to the request for quote list item
            requestForQuoteListItem.SystemUpdate();
        }

        /// <summary>
        /// If the Order Item manufacturer part number & quantity values match a record from within the Quotations History list, use the history list item
        /// to populate certain properties w/ default values based on the history data.
        /// </summary>
        private void LookupOrderItemHistoryData(OrderItemBase orderItem)
        {
            try
            {
                string manufacturerPartNum = orderItem.ManufacturerPartNumber;
                string quantity = orderItem.Quantity;

                SPListItem rfqHistListItem = FindRFQHistoryListItem(manufacturerPartNum, quantity); // lookup a match ...
                if (rfqHistListItem != null) // if found, retrieve desired historical values for the Order Item:
                {
                    // Grab history list fields:
                    string histValue = rfqHistListItem["Internal Sales Rep"].ToString();
                    string vendorUnitPrice = rfqHistListItem["Unit Vendor Price"].ToString();
                    string berryCompliant = rfqHistListItem["Berry Compliant"].ToString();
                    string countryOfOrigin = rfqHistListItem["Country of Origin"].ToString();
                    string minSellPrice = rfqHistListItem["Minimum Sell Price"].ToString();
                    string dealStrategy = rfqHistListItem["Deal Strategy"].ToString();
                    string vendorCost = rfqHistListItem["Vendor Cost"].ToString();

                    // Apply history list fields to order item properties:
                    orderItem.VendorUnitPrice = (Double.Parse(vendorUnitPrice) / 1.054).ToString(); // back out contract fee (1.054)
                    orderItem.BerryAmendmentCompliant = berryCompliant.ToUpper();
                    //orderItem.CountryOfOrigin = countryOfOrigin; /* ***NOTE: GOING TO COMMENT OUT AS THIS IS SUPPOSED TO BE A 3-CHAR FIELD, BUT THE HISTORY LIST VALUE IS THE FULL COUNTRY NAME.*** */
                    orderItem.MinimumSellPrice = minSellPrice;
                    orderItem.DealStrategy = dealStrategy;
                    orderItem.ADSCost = vendorCost;
                }
            }
            catch (Exception ex)
            {
                string logMsg = "Error occured in history list lookup for Order Item: " + orderItem.OrderItemID.ToString() + ". Inner Exception: " + ex.Message.ToString();
                Common.LogEvent(this.OperationListItem.Web, RFQ_EVENT_LOGGER_LIST, this.RFQQuoteNumber, logMsg);
            }
        }

        /// <summary>
        /// Parse's the new Request For Quote file into the appropriate custom SharePoint list.
        /// </summary>
        private void PopulateRFQOrderItemsList()
        {
            // Variables
            SPList rfqActiveList = this.RFQOrderItemsActiveList;
            SPList rfqArchiveList = this.RFQOrderItemsArchiveList;
            string orderItemId;
            string quoteNum;            
            List<OrderItemBase> orderItems = this.OrderItems;

            try
            {
                foreach (var orderItem in orderItems)
                {
                    orderItemId = orderItem.OrderItemID;
                    quoteNum = orderItem.RFQQuoteNumber;

                    // Make sure this Order Item does not already exist within either the "Active" or "Archive" lists:
                    if (!OrderItemAlreadyExists(rfqActiveList, orderItemId, quoteNum) &&
                        (!OrderItemAlreadyExists(rfqArchiveList, orderItemId, quoteNum)))
                    {
                        // Start new list item 
                        SPListItem newListItem = rfqActiveList.Items.Add();

                        // Use Reflection to get property information for the Order Items
                        PropertyInfo[] props = orderItem.GetType().GetProperties();

                        // Loop through & set the values for each column of the new list item using the property values of the current Order Item:
                        string propNm = "";
                        foreach (PropertyInfo prop in props)
                        {
                            // The name of the SharePoint list field that this property targets is stored as an Attribute value. Retrieve the custom attribute:
                            Entities.ColumnAttributes attrs =
                                (Entities.ColumnAttributes)Attribute.GetCustomAttribute(prop, typeof(Entities.ColumnAttributes), false);
                            propNm = attrs.SPColumnName; // grab the SharePoint column name string
                            var propValue = prop.GetValue(orderItem, null); // get this property's value for the current order item object

                            if (propValue != null) // assign non-null property values:
                            {
                                try
                                {
                                    newListItem[propNm] = propValue; // attempt to assign the property value to the new list item using the property name to match against a field on the list item
                                }
                                catch (FormatException formatEx)
                                {
                                    if (formatEx.Message.Contains("DateTime"))
                                    {
                                        newListItem[propNm] = ExcelDocumentUtil.ReadExcelDateTimeValue(propValue); // set DateTime value
                                    }
                                }
                                catch (ArgumentException argEx)
                                {
                                    // Store custom exception error message:
                                    string errMsg = argEx.Message.ToString();
                                    string logDetails = string.Format("Error accessing field in list '{0}': {1}\n" +
                                        "Order Item ID: {2}\n" +
                                        "Field: {3}\n" +
                                        "Value: {4}", rfqActiveList.Title.ToString(), errMsg, orderItemId, propNm, propValue);

                                    // Throw an RFQException w/ full details:
                                    throw new RFQException(logDetails, quoteNum);
                                }
                            }

                        } // end foreach property info

                        // Save changes to list
                        newListItem.Update();
                    }

                } // end foreach order item
            }
            catch (Exception ex)
            {
                throw new Exception("Error occured adding RFQ Order Items to " + rfqActiveList.Title + " list. Inner Exception: " + ex.Message);
            }
        }
                
        /// <summary>
        /// Performs a lookup on the Quotation History list to determine if a Sales Rep can be assigned to this RFQ by default.
        /// </summary>
        /// <returns>The name of the Sales Rep to assign by default if the conditions are met.</returns>
        private string LookupSalesRep()
        {
            // If the RFQ manufacturer part number matches a record from within the RFQ History list, this will associate a Sales Rep
            // that should automatically be set for the RFQ. If no existing Sales Rep is found, the RFQ status will default to 'Unassigned'.
            List<OrderItemBase> orderItems = this.OrderItems;
            string manufacturerPartNum = "";
            string quantity = "";
            string salesRep = "";

            try
            {
                foreach (var orderItem in orderItems)
                {
                    // For each order item, use its manufacturer part number & quantity values to do a lookup on the history list in SharePoint
                    manufacturerPartNum = orderItem.ManufacturerPartNumber;
                    quantity = orderItem.Quantity;

                    SPListItem rfqHistListItem = FindRFQHistoryListItem(manufacturerPartNum, quantity); // lookup a match ...
                    if (rfqHistListItem != null) // if found, retrieve sales rep value for the Request for Quote library:
                    {
                        string histValue = rfqHistListItem["Internal Sales Rep"].ToString();
                        if (!String.IsNullOrEmpty(salesRep) && salesRep != histValue) // if 2 or more matches are found & the sales rep value is DIFFERENT than the first ... 
                        {
                            salesRep = ""; // ... we remove the sales rep value so the RFQ will default to the Unassigned status
                            break;
                        }
                        else
                        {
                            salesRep = histValue;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error occured in history list lookup for Sales Rep. Inner Exception: " + ex.Message);
            }

            return salesRep; // return name of Sales Rep
        }        

        /// <summary>
        /// Determine the default status of the RFQ based upon whether a Sales Rep was found to be automatically set.
        /// </summary>
        /// <returns>The status this RFQ should default to.</returns>
        private string DetermineRFQStatus()
        {
            // If a Sales Rep was assigned by default, set status to Received
            if (!String.IsNullOrEmpty(this.SalesRep))
            {
                return RFQStatusTypes.RECEIVED;
            }
            else // else keep status at Unassigned
            {
                return RFQStatusTypes.UNASSIGNED;
            }
        }

        /// <summary>
        /// Parses the RFQ's filename to determine the contract type.
        /// </summary>
        /// <returns>String representing the contract type of the RFQ.</returns>
        private string DetermineRFQContractType()
        {
            string listItemNm = this.OperationListItem["Name"].ToString().ToUpper();
            string contractType = "";

            // The contract type value is found within the list item's name
            if (listItemNm.Contains(ContractType.TENT.ToString()))
            {
                contractType = ContractType.TENT.ToString();
            }
            else if (listItemNm.Contains(ContractType.SOE.ToString()))
            {
                contractType = ContractType.SOE.ToString();
            }
            else if (listItemNm.Contains(ContractType.FES.ToString()))
            {
                contractType = ContractType.FES.ToString();
            }

            // Return contract type string
            return contractType;
        }

        /// <summary>
        /// Determine which "Active" SharePoint List to target depending on the rfq's contract type.
        /// </summary>
        /// <param name="web">The SharePoint Web context from which to pull the list.</param>
        /// <returns>The SharePoint List, if it was found.</returns>
        private SPList GetTargetOrderItemsActiveList(SPWeb web)
        {
            string contractType = this.ContractNm;
            string listName = contractType; // determine list name based on contract type
            SPList targetList = null;

            // Attempt to retrieve a SharePoint list w/ the specified list name
            targetList = web.Lists.TryGetList(listName);

            return targetList; // return correct list for Request for Quote's contract type
        }

        /// <summary>
        /// Determine which "Archive" SharePoint List to target depending on the rfq's contract type.
        /// </summary>
        /// <param name="web">The SharePoint Web context from which to pull the list.</param>
        /// <returns>The Order Item archive SharePoint List, if it was found.</returns>
        private SPList GetTargetOrderItemsArchiveList(SPWeb web)
        {
            string contractType = this.ContractNm;
            string listName = contractType + " Archive";

            return web.Lists.TryGetList(listName); // attempt to retrieve a SharePoint list w/ the specified list name           
        }

    }
}
