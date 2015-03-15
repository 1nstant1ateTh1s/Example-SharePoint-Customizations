using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;

namespace RFQEventReceiver
{
    public class Common
    {
        /// <summary>
        /// Given a list item, will start the specified associated workflow.
        /// </summary>
        /// <param name="listItem">The list item to call this workflow for.</param>
        /// <param name="workflowName">The name of the workflow to start.</param>
        public static void StartWorkflow(SPListItem listItem, string workflowName)
        {
            SPWorkflowManager wfMgr = listItem.Web.Site.WorkflowManager;
            SPWorkflowAssociation wfAssoc = listItem.ParentList.WorkflowAssociations.GetAssociationByName(workflowName, 
                System.Globalization.CultureInfo.CurrentCulture);

            if (wfAssoc != null)
            {
                // Use the Workflow Manager to start the specified workflow
                wfMgr.StartWorkflow(listItem, wfAssoc, wfAssoc.AssociationData, true);
            }
        }

        /// <summary>
        /// Log the event to the specified list.
        /// </summary>
        /// <param name="web">The current web context.</param>
        /// <param name="listName">The name of the Logs list to target.</param>
        /// <param name="eventType">The type of event that occured.</param>
        /// <param name="rfqQuoteNum"></param>
        /// <param name="details">The details of the event to log.</param>
        public static void LogEvent(SPWeb web, string listName, SPEventReceiverType eventType, string rfqQuoteNum, string details)
        {
            // Retrieve reference to Log list
            SPList logList = Common.EnsureLogList(web, listName);

            // Add item to Log list
            SPListItem logItem = logList.Items.Add();
            logItem["RFQ#"] = rfqQuoteNum;
            logItem["Title"] = string.Format("{0} triggered at {1}", eventType, DateTime.Now);
            logItem["Event"] = eventType.ToString();
            logItem["Date"] = DateTime.Now;
            logItem["Details"] = details;
            logItem.Update(); // save list item
        }

        /// <summary>
        /// Log the event to the specified list.
        /// </summary>
        /// <param name="web">The current web context.</param>
        /// <param name="listName">The name of the Logs list to target.</param>
        /// <param name="rfqQuoteNum"></param>
        /// <param name="details">The details of the event to log.</param>
        public static void LogEvent(SPWeb web, string listName, string rfqQuoteNum, string details)
        {
            // Retrieve reference to Log list
            SPList logList = Common.EnsureLogList(web, listName);

            // Add item to Log list
            SPListItem logItem = logList.Items.Add();
            logItem["RFQ#"] = rfqQuoteNum;
            logItem["Title"] = string.Format("Message logged at {0}", DateTime.Now);
            logItem["Date"] = DateTime.Now;
            logItem["Details"] = details;
            logItem.Update(); // save list item
        }

        /// <summary>
        /// Determines whether a log list already exists. If it does, we get a reference to it. Otherwise,
        /// we create the list & return a reference to it.
        /// </summary>
        /// <param name="web">The current web context.</param>
        /// <param name="listName">The name of the Logs list to target.</param>
        /// <returns>A reference to the Logs list with the specified name.</returns>
        private static SPList EnsureLogList(SPWeb web, string listName)
        {
            SPList list = null;
            try
            {
                list = web.Lists[listName]; // attempt to retrieve a list w/ the given name
            }
            catch
            {
                // List does not exist, so we need to create a new list
                Guid listGuid = web.Lists.Add(listName, listName, SPListTemplateType.GenericList);
                list = web.Lists[listGuid];
                list.OnQuickLaunch = true;

                // Add the fields to the list
                // No need to add "Title" since it's already added by default. We use it to set the event name
                list.Fields.Add("RFQ#", SPFieldType.Text, true);
                list.Fields.Add("Event", SPFieldType.Text, true);
                list.Fields.Add("Date", SPFieldType.DateTime, true);
                list.Fields.Add("Details", SPFieldType.Note, false);

                // Specify what fields to view
                SPView view = list.DefaultView;
                view.ViewFields.Add("RFQ#");
                view.ViewFields.Add("Event");
                view.ViewFields.Add("Date");
                view.ViewFields.Add("Details");
                view.Update(); // save view changes

                list.Update(); // save list changes
            }

            return list;
        }
    }
}
