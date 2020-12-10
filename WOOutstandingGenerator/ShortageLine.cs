using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WOOutstandingGenerator
{
    class ShortageLine
    {
        public string PartNo { get; set; }
        public string Description { get; set; }
        public string WorksOrderNumber { get; set; }
        public string WOProductGroupCode { get; set; }
        public Nullable<System.DateTime> WODueDate { get; set; }
        public string WORaisedBy { get; set; }
        public string ProductGroupCode { get; set; }
        public string PartMethod { get; set; }
        public string Responsibility { get; set; }
        public string Supplier { get; set; }
        public string CommercialNotes { get; set; }
        public string BatchNotes { get; set; }
        public Nullable<decimal> Demand { get; set; }
        public Nullable<decimal> DemandForThisDate { get; set; }
        public Nullable<decimal> GoodStock { get; set; }
        public Nullable<decimal> BadStock { get; set; }
        public decimal NetShortage { get; set; }
        public Nullable<decimal> StockLeftAfterThisDate { get; set; }
        public string PurchaseOrderNumber { get; set; }
        public Nullable<System.DateTime> PurchaseOrderDeliveryDate { get; set; }
        public Nullable<decimal> PurchaseOrderQty { get; set; }
        public string PORaisedBy { get; set; }
        public string ComponentWorksOrder { get; set; }
        public Nullable<System.DateTime> ComponentWODueDate { get; set; }
        public Nullable<decimal> ComponentWOQuantity { get; set; }
        public string ParentAssembly { get; set; }
        public string ParentAssemblyDescription { get; set; }
        public bool Issued { get; set; }
        public bool POCoversDemand { get; set; }
        public decimal UnitCost { get; set; }
        public decimal Store1 { get; set; }
        public decimal Store2 { get; set; }
        public decimal Store3 { get; set; }
        public decimal Store4 { get; set; }
        public decimal MoyFab { get; set; }
        public decimal EagleOverseas { get; set; }
        public string GoodLocations { get; set; }
        public string BadLocations { get; set; }
        public string compResponsibility { get; set; }
        public bool IsStoresRequest { get; set; }

        

    }
}
