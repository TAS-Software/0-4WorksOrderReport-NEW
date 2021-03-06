//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WOOutstandingGenerator
{
    using System;
    using System.Collections.Generic;
    
    public partial class WOLineReport
    {
        public long WOLine_ID { get; set; }
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
        public string PORaisedBy { get; set; }
        public string WORaisedBy { get; set; }
        public string WORespCode { get; set; }
        public Nullable<decimal> UnitCost { get; set; }
        public string POComments { get; set; }
        public string CompRespCode { get; set; }
        public Nullable<bool> IsStoresRequest { get; set; }
        public Nullable<decimal> SO_Demand { get; set; }

        public Nullable<decimal> DemandForThisDate { get; set; }

        public bool POCoversDemand { get; set; }

        public string Owner { get; set; }

        public decimal Store1 { get; set; }
        public decimal Store2 { get; set; }
        public decimal Store3 { get; set; }
        public decimal Store4 { get; set; }
        public decimal MoyFab { get; set; }
        public decimal EagleOverseas { get; set; }
        public string GoodLocations { get; set; }
        public string BadLocations { get; set; }
    }
}
