
var rfqObj = rfqObj || {};

/* js object for generating rfq return file */
rfqObj.rfqReturnFile = {
    camlQueries: {
        GET_BY_ID_QUERY_STRING: function (id) {
            return ('<View><Query><Where><Eq>' +
					'<FieldRef Name=\'ID\'/>' +
					'<Value Type=\'Integer\'>' + id + '</Value></Eq></Where>' +
					'</Query></View>');
        },
        RFQ_ORDER_ITEMS_QUERY_STRING: function (rfqNum) {
            return ('<View><Query><Where><Eq>' +
					'<FieldRef Name=\'Quote_Number\'/>' +
					'<Value Type=\'Text\'>' + rfqNum + '</Value></Eq></Where>' +
					'</Query></View>');
        }
    },
    contractTypeInputFields: {
        TENT: 'Title,Quote_Number,TLSP_Vendor_Part_Number,Transportation_Price,Vendor_Unit_Price,Purchase_Extended_Price,Lead_Time,Comments,Procurement_Agreement_Compliant,' +
                 'Berry_Amendment_Compliant,Country_of_Origin,Alt_Core_List_Number,Alt_NSN,Alt_Manufacturer_Name,Alt_Manufacturer_Part_Number,' +
                 'Alt_TLSP_Vendor_Part_Number,Alt_Item_Description,Alt_Transportation_Price,Alt_Vendor_Unit_Price,Alt_Lead_Time,Alt_Comments,' +
                 'Alt_PA_Compliant,Alt_Berry_Amendment_Compliant,Alt_Country_of_Origin',
        SOE: 'Title,Quote_Number,TLSP_Vendor_Part_Number,Transportation_Price,Vendor_Unit_Price,Purchase_Extended_Price,Lead_Time,Comments,Procurement_Agreement_Compliant,' +
                 'Berry_Amendment_Compliant,Country_of_Origin,Alt_Core_List_Number,Alt_NSN,Alt_Manufacturer_Name,Alt_Manufacturer_Part_Number,' +
                 'Alt_TLSP_Vendor_Part_Number,Alt_Item_Description,Alt_Transportation_Price,Alt_Vendor_Unit_Price,Alt_Lead_Time,Alt_Comments,' +
                 'Alt_PA_Compliant,Alt_Berry_Amendment_Compliant,Alt_Country_of_Origin',
        FES: 'Title,Quote_Number,TLSP_Vendor_Part_Number,Transportation_Price,Vendor_Unit_Price,Purchase_Extended_Price,Lead_Time,Comments,Trade_Agreement_Compliant,' +
                'Berry_Amendment_Compliant,Country_of_Origin,Alt_Core_List_Number,Alt_NSN,Alt_Manufacturer_Name,Alt_Manufacturer_Part_Number,' +
                'Alt_TLSP_Vendor_Part_Number,Alt_Item_Description,Alt_Transportation_Price,Alt_Vendor_Unit_Price,Alt_Lead_Time,Alt_Comments,' +
                'Alt_Trade_Agreement_Compliant,Alt_Berry_Amendment_Compliant,Alt_Country_of_Origin',
        getOrderItemFields: function () {
            /// <summary></summary>
            /// <returns></returns>
            var currentCtx = GetCurrentCtx(), // retrieve a ContextInfo object w/ some useful properties
                fields = null,
                listTitle = "";

            if (currentCtx.ListTitle != undefined) {
                listTitle = currentCtx.ListTitle; // use the ContextInfo object to determine the currently selected List title

                if (listTitle.indexOf("SOE") >= 0) {
                    fields = this.SOE;
                }
                else if (listTitle.indexOf("FES") >= 0) {
                    fields = this.FES;
                }
                else if (listTitle.indexOf("TENT") >= 0) {
                    fields = this.TENT;
                }
            }
            return fields;
        }
    },

    _rfqList: null,
    _selListItem: null,
    _rfqOrderItems: null,
    _context: null,
    getRfqList: function () {
        /// <summary>Accessor for the _rfqList property.</summary>
        /// <returns></returns>
        return this._rfqList;
    },
    setRfqList: function (value) {
        /// <summary>Mutator function for the _rfqList property.</summary>
        /// <param name="value"></param>
        if (value != undefined) {
            this._rfqList = this.getContext().get_web().get_lists().getById(value);
        }
    },
    getSelListItem: function () {
        /// <summary>Accessor function for the _selListItem property.</summary>
        /// <returns></returns>
        return this._selListItem;
    },
    setSelListItem: function (value) {
        /// <summary>Mutator function for the _selListItem property.</summary>
        /// <param name="value"></param>
        if (value != undefined) {
            this._selListItem = value;
        }
    },
    getRfqOrderItems: function () {
        /// <summary>Accessor for the _rfqOrderItems property.</summary>
        /// <returns></returns>
        return this._rfqOrderItems;
    },
    setRfqOrderItems: function (value) {
        /// <summary>Mutator function for the _selListItem property.</summary>
        /// <param name="value"></param>
        if (value != undefined) {
            this._rfqOrderItems = value;
        }
    },
    getContext: function () {
        /// <summary>Accessor for the _context property.</summary>
        if (this._context == null) {
            this._context = new SP.ClientContext.get_current(); // get the current client context
        }
        return this._context;
    },

    isEnabled: function () {
        /// <summary></summary>
        /// <returns></returns>
        var items = SP.ListOperation.Selection.getSelectedItems();
        return (items.length == 1);
    },
    generate: function (listId, itemId) {
        /// <summary></summary>
        /// <param name="listId"></param>
        /// <param name="itemId"></param>
        // variables
        var selectedListId = listId,
			selectedItemId = itemId;

        // set currently selected rfq list
        this.setRfqList(selectedListId);

        // determine the rfq quote number of the currently selected row in the rfq list
        this.getSelectedQuoteNum(selectedItemId);
    },
    getSelectedQuoteNum: function (id) {
        /// <summary>Gets the RFQ Quote Number of the currently selected list item.</summary>
        /// <param name="id" type="string">The ID of the list item from which to retrieve the rfq quote number.</param>
        var context = this.getContext(),
            rfqList = this.getRfqList(),
            camlQuery = new SP.CamlQuery();

        // attach query string to query object
        camlQuery.set_viewXml(this.camlQueries.GET_BY_ID_QUERY_STRING(id));
        this.setSelListItem(rfqList.getItems(camlQuery));

        // declare what to load (the 'Quote_Number' field, which displays the RFQ Quote Number)
        context.load(this.getSelListItem(),
            'Include(Quote_Number)');

        // run the query on the server
        context.executeQueryAsync(Function.createDelegate(this, this.quoteNumRetrievedSuccess),
			Function.createDelegate(this, this.onFailedQuery));
    },
    getSelectedRfqOrderItems: function (quoteNum) {
        /// <summary>Gets the order items in the list that belong to the selected RFQ.</summary>
        /// <param name="quoteNum" type="string">The RFQ Quote Number on which to search for List Items.</param>
        var context = this.getContext(),
            rfqList = this.getRfqList(),
            camlQuery = new SP.CamlQuery(),
            orderItemFields = this.contractTypeInputFields.getOrderItemFields();

        // attach query string to query object
        camlQuery.set_viewXml(this.camlQueries.RFQ_ORDER_ITEMS_QUERY_STRING(quoteNum));
        this.setRfqOrderItems(rfqList.getItems(camlQuery));

        // declare what to load
        context.load(this.getRfqOrderItems(),
            'Include(' + orderItemFields + ')'
        );

        // run the query
        context.executeQueryAsync(Function.createDelegate(this, this.rfqOrderItemsRetrievedSuccess),
            Function.createDelegate(this, this.onFailedQuery));
    },

    quoteNumRetrievedSuccess: function (sender, args) {
        /// <summary>Handles the successful retrieval of the RFQ Quote Number.</summary>
        var listItems = this.getSelListItem(),
            listItemEnumerator,
            rfqQuoteNum = "";

        // check that items exist
        if (listItems != null) {
            listItemEnumerator = listItems.getEnumerator(); // get enumerator
            while (listItemEnumerator.moveNext()) { // move to item
                // retrieve current item's "Quote_Number" field (i.e., the RFQ Quote Number):
                var item = listItemEnumerator.get_current();
                rfqQuoteNum = item.get_item('Quote_Number');
            }
            // use this quote number value to retrieve all order items associated to the rfq
            this.getSelectedRfqOrderItems(rfqQuoteNum);
        }
    },
    rfqOrderItemsRetrievedSuccess: function (sender, args) {
        /// <summary>Handles the successful retrieval of all order items associated with the current RFQ.</summary>
        var listItems = this.getRfqOrderItems(),
            listItemEnumerator,
            itemObj = {},
            items = [];

        // check that items exist
        if (listItems != null) {
            listItemEnumerator = listItems.getEnumerator(); // get enumerator
            while (listItemEnumerator.moveNext()) { // enumerator over collection
                // retrieve current item's field collection & add to array:
                itemObj = listItemEnumerator.get_current().get_fieldValues();
                items.push(itemObj); // add object to array
            }

            // call web service w/ ajax, passing array of field values from each order item in this rfq collection, to
            // perform the interaction w/ Excel.
            var svcNm = "GenRFQReturnFile" + GetCurrentCtx().ListTitle;
            this.callService(svcNm, items);
        }
    },
    onFailedQuery: function (sender, args) {
        /// <summary>Action to take on query failure.</summary>
        SP.UI.Notify.addNotification("Query failed: " + args.get_message(), false);
    },

    callService: function (svcNm, parms) {
        /// <summary></summary>
        /// <param name="svcNm"></param>
        /// <param name="parms"></param>
        var svcUrl = "/_layouts/RfqWorkflow/WebServices/RfqWorkflowService.svc/" + svcNm, // build url w/ supplied service method name
            relativeUrl = this.getContext().get_url(),
            parmstring = (parms != undefined ? JSON.stringify({ orderItems: parms, relativeUrl: relativeUrl }) : ""); // default to empty string if no parameters were supplied for 
        // service call; otherwise, stringify object of parameters
        // call the web service code
        $.ajax({
            type: "POST",
            contentType: "application/json",
            dataType: "json",
            url: svcUrl,
            data: parmstring,
            success: function (msg) {
                var returnObj = (msg[svcNm + 'Result'] != undefined ? msg[svcNm + 'Result'] : {}),
                    errors = (returnObj.Errors != undefined ? returnObj.Errors : []),
                    errCount = errors.length,
                    i = 0,
                    displayMsg = "";

                if (errCount > 0) { // indicates error
                    displayMsg = "Error executing Generate RFQ Return File service: ";
                    for (i = 0; i < errCount; i++) { // append each error message to be displayed
                        displayMsg += "<br />" + errors[i].toString();
                    }
                }
                else { // indicates success
                    displayMsg = "Generate RFQ Return File Success";
                }

                // display notification
                SP.UI.Notify.addNotification(displayMsg);
            },
            error: function (msg, url, status) {
                SP.UI.Notify.addNotification("Error: " + status, false);
            }
        });
    }
};

/* js object for submitting rfq bid */
rfqObj.submitRfqBid = {

    /* TO-DO: TRY TO TURN INTO A FUNCTION THAT WILL FIRST LOAD/REQUIRE AN EXTERNAL .JS FILE & 
    THEN RETURN THIS "OBJECT", TO SEPERATE THIS CODE OUT. */

    process: function (listId, listItemId) {
        /// <summary></summary>
        /// <param name="listId"></param>
        /// <param name="listItemId"></param>
        var context = SP.ClientContext.get_current(),
            list = context.get_web().get_lists().getById(listId),
            selListItem = list.getItemById(listItemId);

        // declare what to load
        context.load(selListItem,
            'ID',
            'Status'
        );

        // run the query on the server
        context.executeQueryAsync(Function.createDelegate(selListItem, this.statusRetrievedSuccess),
			Function.createDelegate(this, this.onFailedQuery));
    },
    statusRetrievedSuccess: function (sender, args) {
        /// <summary></summary>

        var item = this.get_fieldValues(),
            status = (item.Status != null ? item.Status : "");

        // if the RFQ item is ready to be sent, then submit the RFQ bid via e-mail
        if (status == "Approved") {

            // call web service to submit bid (i.e., e-mail Excel document, update Status to "Submitted", etc.)
            var url = "/_layouts/RfqWorkflow/WebServices/RfqWorkflowService.svc/SubmitRFQBid",
                relativeUrl = SP.ClientContext.get_current().get_url(),
                parmstring = JSON.stringify({ rfqItemId: item.ID, relativeUrl: relativeUrl });

            // call the web service code
            $.ajax({
                type: "POST",
                contentType: "application/json",
                dataType: "json",
                url: url,
                data: parmstring,
                success: function (msg) {
                    var returnObj = (msg.SubmitRFQBidResult != null ? msg.SubmitRFQBidResult : {}),
                        errors = (returnObj.Errors != undefined ? returnObj.Errors : []),
                        errCount = errors.length,
                        i = 0,
                        displayMsg = "";

                    if (errCount > 0) { // indicates error
                        displayMsg = "Error Submitting RFQ Bid: ";
                        for (i = 0; i < errCount; i++) { // append each error message to be displayed
                            displayMsg += "<br />" + errors[i].toString();
                        }
                    }
                    else { // indicates success
                        displayMsg = "RFQ Bid Submitted Successfully";
                    }

                    // display notification
                    SP.UI.Notify.addNotification(displayMsg);

                },
                error: function (msg, url, status) {
                    SP.UI.Notify.addNotification("Error Submitting RFQ Bid: " + status, false);
                }
            });

        }
        else {
            SP.UI.Notify.addNotification("RFQ bid not ready for submission. RFQ must be in the 'Approved' status.");
        }
    },
    onFailedQuery: function (sender, args) {
        /// <summary>Action to take on query failure.</summary>
        SP.UI.Notify.addNotification("Query failed: " + args.get_message(), false);
    },
    isEnabled: function () {
        /// <summary></summary>
        /// <returns></returns>
        var items = SP.ListOperation.Selection.getSelectedItems();
        return (items.length == 1);
    }
};