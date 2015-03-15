using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace RFQ_SharePoint_Project.WebServices
{
    [ServiceContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public interface IRFQWorkflowService
    {
        [OperationContract]
        [WebInvoke(UriTemplate = "/GenRFQReturnFileSOE", Method = "POST", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        JsonResponse GenRFQReturnFileSOE(SOE_RFQOrderItems orderItems, string relativeUrl);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GenRFQReturnFileFES", Method = "POST", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        JsonResponse GenRFQReturnFileFES(FES_RFQOrderItems orderItems, string relativeUrl);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GenRFQReturnFileTENT", Method = "POST", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        JsonResponse GenRFQReturnFileTENT(TENT_RFQOrderItems orderItems, string relativeUrl);

        [OperationContract]
        [WebInvoke(UriTemplate = "/SubmitRFQBid", Method = "POST", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        JsonResponse SubmitRFQBid(int rfqItemId, string relativeUrl);
    }

    [CollectionDataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class RFQOrderItems<T> : List<T>
    {
    }

    [CollectionDataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class SOE_RFQOrderItems : RFQOrderItems<SOE_RFQOrderItem>
    {
    }

    [CollectionDataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class FES_RFQOrderItems : RFQOrderItems<FES_RFQOrderItem>
    {
    }

    [CollectionDataContract(Namespace = "RFQ_SharePoint_Project.WebServices.RFQWorkflowService")]
    public class TENT_RFQOrderItems : RFQOrderItems<TENT_RFQOrderItem>
    {
    }

    [DataContract]
    public class JsonResponse
    {
        [DataMember]
        public bool IsSuccess { get; set; }
        [DataMember]
        public List<string> Errors = new List<string>();
    }
}
