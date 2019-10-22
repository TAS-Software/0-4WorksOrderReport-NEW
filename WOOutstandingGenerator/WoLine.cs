using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WOOutstandingGenerator
{
    class WoLine
    {
        public string WorksOrderNumber { get; set; }
        public string ParentStatusCode { get; set; }
        public string ProductGroupCode { get; set; }
        public string ParentPartNumber { get; set; }
        public string ParentPartDescription { get; set; }
        public string ParentMethod { get; set; }
        public Nullable<decimal> OrderQuantity { get; set; }
        public Nullable<decimal> WOOutstanding { get; set; }
        public string ComponentPartNumber { get; set; }
        public string ComponentPartDescription { get; set; }
        public string ComponentStatusCode { get; set; }
        public string CurrentComponentMethodType { get; set; }
        public string ComponentCurrentRev { get; set; }
        public Nullable<decimal> PlannedIssueQuantity { get; set; }
        public Nullable<System.DateTime> PlannedIssueDate { get; set; }
        public Nullable<System.DateTime> CompletionDate { get; set; }
        public Nullable<decimal> ActualIssueQuantity { get; set; }
        public Nullable<System.DateTime> ActualIssueDate { get; set; }
        public Nullable<decimal> ReturnQuantity { get; set; }
        public Nullable<decimal> Outstanding { get; set; }
        public Nullable<decimal> DemandUptoThisPlannedIssueDate { get; set; }
        public Nullable<decimal> StockLeftAfterThisDate { get; set; }
        public Nullable<decimal> GoodStock { get; set; }
        public Nullable<decimal> BadStock { get; set; }
        public string ComponentGroupCode { get; set; }
        public string WOCommercialNotes { get; set; }
        public string WOProductionNotes { get; set; }
        public string PurchaseOrderNumber { get; set; }
        public Nullable<decimal> QuantityPurchased { get; set; }
        public Nullable<System.DateTime> ReceiptDate { get; set; }
        public string SupplierName { get; set; }
        public string POPartNotes { get; set; }
        public string ComponentWorksOrder { get; set; }
        public Nullable<System.DateTime> WODueDate { get; set; }
        public Nullable<decimal> Quantity { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public bool Issued { get; set; }
        public bool POCoversDemand { get; set; }
        public string compResponsibility { get; set; }

    }
}
