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
    
    public partial class WOLineReport_WOPartsLevel
    {
        public long WOLine_ID { get; set; }
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string WONumber { get; set; }
        public Nullable<System.DateTime> WODueDate { get; set; }
        public string ProductGroup { get; set; }
        public string PartMethod { get; set; }
        public string Responsibility { get; set; }
        public string CommercialNotes { get; set; }
        public string BatchNotes { get; set; }
        public Nullable<decimal> Demand { get; set; }
        public Nullable<decimal> DemandForThisDate { get; set; }
        public Nullable<decimal> GoodStock { get; set; }
        public Nullable<decimal> BadStock { get; set; }
        public Nullable<decimal> NetShortage { get; set; }
        public Nullable<decimal> StockAfterThisDate { get; set; }
        public string Supplier { get; set; }
        public string PONumber { get; set; }
        public Nullable<System.DateTime> PODeliveryDate { get; set; }
        public Nullable<decimal> POQuantity { get; set; }
        public string PORaisedBy { get; set; }
        public string ComponentWO { get; set; }
        public Nullable<System.DateTime> ComponentWODueDate { get; set; }
        public Nullable<decimal> ComponentWOQuantity { get; set; }
        public string WORaisedBy { get; set; }
        public string ParentAssembly { get; set; }
        public string ParentAssemblyDescription { get; set; }
        public bool Issued { get; set; }
        public Nullable<bool> POCoversDemand { get; set; }
        public string Owner { get; set; }
        public Nullable<decimal> UnitCost { get; set; }
        public Nullable<decimal> Store1 { get; set; }
        public Nullable<decimal> Store2 { get; set; }
        public Nullable<decimal> Store3 { get; set; }
        public string OtherGood { get; set; }
        public string OtherBad { get; set; }
        public string AllWOs { get; set; }
        public string POComments { get; set; }
        public string CompRespCode { get; set; }
        public Nullable<bool> StoresRequest { get; set; }
    }
}
