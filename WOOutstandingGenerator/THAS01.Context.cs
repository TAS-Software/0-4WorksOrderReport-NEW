﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Entity.Core.Objects;
    using System.Linq;
    
    public partial class thas01ReportEntities : DbContext
    {
        public thas01ReportEntities()
            : base("name=thas01ReportEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
    
        public virtual ObjectResult<THAS_CONNECT_StockLocationCount_Result> THAS_CONNECT_StockLocationCount()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<THAS_CONNECT_StockLocationCount_Result>("THAS_CONNECT_StockLocationCount");
        }
    
        public virtual int WODumpProcedure()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("WODumpProcedure");
        }
    
        public virtual int THAS_CONNECT_OPENWO_NEW_V2()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("THAS_CONNECT_OPENWO_NEW_V2");
        }
    
        public virtual ObjectResult<THAS_CONNECT_OnlineAvailable_Result> THAS_CONNECT_OnlineAvailable()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<THAS_CONNECT_OnlineAvailable_Result>("THAS_CONNECT_OnlineAvailable");
        }
    
        public virtual ObjectResult<THAS_CONNECT_OnlineShortage_Result> THAS_CONNECT_OnlineShortage()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<THAS_CONNECT_OnlineShortage_Result>("THAS_CONNECT_OnlineShortage");
        }
    }
}
