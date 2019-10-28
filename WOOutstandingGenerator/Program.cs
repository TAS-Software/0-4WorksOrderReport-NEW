using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WOOutstandingGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "In Build Shortage Report Generator [" + typeof(WOOutstandingGenerator.Program).Assembly.GetName().Version + "]";
            
            TimeSpan _timeToRun = new TimeSpan(05, 30, 00);
            TimeSpan _closeTime = new TimeSpan(23, 00, 00);
            Console.WriteLine("Do you want to run now first?");
            var answer = Console.ReadKey();

            if (answer.KeyChar == 'y')
            {
                ProcessShortageReport(); //Taken Out For Now
                while (true)
                {

                    TimeSpan timeNow = DateTime.Now.TimeOfDay;

                    Console.WriteLine(" ----- ");
                    Console.WriteLine("Time Now is : " + timeNow.ToString());
                    if (timeNow > _closeTime || timeNow <= _timeToRun) // cannot run during wee hours of morning before 5.30am as this when MRP runs and BI refreshes
                    {
                        TimeSpan remainsOfDay = new TimeSpan(0, 0, 0);
                        TimeSpan diff = new TimeSpan(0, 0, 0);
                        if (timeNow > _closeTime)
                        {
                            remainsOfDay = new TimeSpan(24, 0, 0) - timeNow;
                            diff = _timeToRun + remainsOfDay;
                        }
                        else
                        {
                            diff = _timeToRun - timeNow;
                        }
                        Console.WriteLine("We are in the deadzone whilst MRP runs.");
                        Console.WriteLine("Time Now is : " + timeNow.ToString());
                        Console.WriteLine("Next Run Time is : " + _timeToRun.ToString());
                        Console.WriteLine("The Wait Window is : " + diff.Duration());
                        //AppLogger.ReportInfo("Time Now is : " + timeNow.ToString());
                        //AppLogger.ReportInfo("Next Run Time is : " + _timeToRun.ToString());
                        System.Threading.Thread.Sleep(diff.Duration());
                    }
                    else  // time to run baby!!!
                    {
                        var h = timeNow.Hours + 1;
                        var newTime = new TimeSpan(h, 15, 00);
                        var wait = (newTime - timeNow).Duration();
                        Console.WriteLine("The Wait Window is : " + wait.ToString());
                        Console.WriteLine("We will sleep for this time.");
                        //AppLogger.ReportInfo("The Wait Window is : " + wait.ToString());
                        //AppLogger.ReportInfo("We will sleep for this time.");
                        System.Threading.Thread.Sleep(wait);
                    }
                    //AppLogger.ReportInfo("We have reached the scheduled run time : " + DateTime.Now.TimeOfDay.ToString());
                    ProcessShortageReport();
                }
            }
            else
            {
                while (true)
                {

                    TimeSpan timeNow = DateTime.Now.TimeOfDay;

                    Console.WriteLine(" ----- ");
                    Console.WriteLine("Time Now is : " + timeNow.ToString());
                    if (timeNow > _closeTime || timeNow <= _timeToRun) // cannot run during wee hours of morning before 6.30am as this when MRP runs and BI refreshes
                    {
                        TimeSpan remainsOfDay = new TimeSpan(0, 0, 0);
                        TimeSpan diff = new TimeSpan(0, 0, 0);
                        if (timeNow > _closeTime)
                        {
                            remainsOfDay = new TimeSpan(24, 0, 0) - timeNow;
                            diff = _timeToRun + remainsOfDay;
                        }
                        else
                        {
                            diff = _timeToRun - timeNow;
                        }
                        Console.WriteLine("We are in the deadzone whilst MRP runs.");
                        Console.WriteLine("Time Now is : " + timeNow.ToString());
                        Console.WriteLine("Next Run Time is : " + _timeToRun.ToString());
                        Console.WriteLine("The Wait Window is : " + diff.Duration());
                        //AppLogger.ReportInfo("Time Now is : " + timeNow.ToString());
                        //AppLogger.ReportInfo("Next Run Time is : " + _timeToRun.ToString());
                        System.Threading.Thread.Sleep(diff.Duration());
                    }
                    else  // time to run baby!!!
                    {
                        var h = timeNow.Hours + 1;
                        var newTime = new TimeSpan(h, 10, 00);
                        var wait = (newTime - timeNow).Duration();
                        Console.WriteLine("The Wait Window is : " + wait.ToString());
                        Console.WriteLine("We will sleep for this time.");
                        //AppLogger.ReportInfo("The Wait Window is : " + wait.ToString());
                        //AppLogger.ReportInfo("We will sleep for this time.");
                        System.Threading.Thread.Sleep(wait);
                    }
                    //AppLogger.ReportInfo("We have reached the scheduled run time : " + DateTime.Now.TimeOfDay.ToString());
                    ProcessShortageReport();
                }
            }
        }

        private static void ProcessShortageReport()
        {
            thas01ReportEntities thas = new thas01ReportEntities();
            ConnectReportDbEntities connect = new ConnectReportDbEntities();
            var owners = connect.BOMShortageProductGroups.Include("BOMShortageOwners").ToList();
            // todo - look at these!!!          
            List<ShortageLineGrouped> exports = new List<ShortageLineGrouped>();
            List<ShortageLine> exports2 = new List<ShortageLine>();
          
            thas.Database.CommandTimeout = 180;
            connect.Database.CommandTimeout = 180;
            string regexPattern = @"\{\*?\\[^{}]+}|[{}]|\\\n?[A-Za-z]+\n?(?:-?\d+)?[ ]?";

            FileInfo fileInfo;
            string theDate = DateTime.Now.ToString("yyyyMMdd");
            string theDateHours = DateTime.Now.ToString("yyyyMMdd HH.mm.ss");
            if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, @"InBuildShortagesReport_", "Shortage Reports", false)) //Shortage Reports
            {
                var wolineTotals = connect.WODumpTotals.ToList();
                List<THAS_CONNECT_StockLocationCount_Result> stockCounts = new List<THAS_CONNECT_StockLocationCount_Result>();

                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    var partsws = excelPackage.Workbook.Worksheets.Add("Parts-Shortages");

                    var resultSet = new List<WOLineReport>();
                    bool succeeded = false;
                    int failCount = 1;
                   
                    while (!succeeded && failCount <= 10)
                    {
                        try
                        {
                            connect.Database.ExecuteSqlCommand("truncate table WODump");
                            connect.Database.ExecuteSqlCommand("truncate table WODumpTotals");
                            connect.Database.ExecuteSqlCommand("truncate table WOLineReport");

                            try
                            {
                                thas.WODumpProcedure();
                                Console.WriteLine("WODump Successful.");

                                try
                                {
                                    thas.THAS_CONNECT_OPENWO_NEW_V2();
                                    succeeded = true;
                                    Console.WriteLine("OpenWO Successful.");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("OpenWO Failed. " + ex.Message);
                                    succeeded = false;
                                    failCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("WODump Failed. " + ex.Message);
                                if (ex.InnerException != null)
                                {
                                    Console.WriteLine("Inner Exception Details: " + ex.InnerException.Message);
                                    succeeded = false;
                                    failCount++;
                                }
                            }

                            using (var rptProd = new thas01ReportEntities())
                            {
                                try
                                {
                                    stockCounts = rptProd.THAS_CONNECT_StockLocationCount().ToList();
                                }
                                catch (Exception)
                                {
                                    throw;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("The Query Failed For The Following Reason: " + ex.Message + ". It Will Attempt To Run Again. Fail Count: " + failCount + ".");
                            failCount++;
                            succeeded = false;
                        }
                    }
                    resultSet = connect.WOLineReports.ToList();
                    var parts = resultSet.Select(x => x.ComponentPartNumber).Distinct().ToList();
                    stockCounts = stockCounts.Where(x => parts.Contains(x.PartNumber)).ToList();
                    Console.WriteLine("The Query Was Successfully Retrieved, Please Wait For The File To Generate.");
                                                  

                    int totalCount = resultSet.ToList().Count; //.Where(x => x.CurrentComponentMethodType == "Purchased").ToList().Count;
                    int ctr = 0;
                    Console.WriteLine("--- " + DateTime.Now + " There are a total of " + totalCount + " lines to process. ---");

                    List<THAS_CONNECT_StockLocationCount_Result> cleaned = new List<THAS_CONNECT_StockLocationCount_Result>();                    

                    foreach (WOLineReport woLine in resultSet)
                    {
                        Regex rgxProd = new Regex(regexPattern);
                        string productionReplace = "";
                        string ProductionNotes = woLine.WOProductionNotes != null ? woLine.WOProductionNotes : "";
                        woLine.WOProductionNotes = rgxProd.Replace(ProductionNotes, productionReplace);

                        Regex rgxComm = new Regex(regexPattern);
                        string CommercialReplace = "";
                        string CommercialNotes = woLine.WOCommercialNotes != null ? woLine.WOCommercialNotes : "";
                        woLine.WOCommercialNotes = rgxComm.Replace(CommercialNotes, CommercialReplace);
                        woLine.WOCommercialNotes = woLine.WOCommercialNotes.TrimEnd(' ');

                        Regex rgxPOComm = new Regex(regexPattern);
                        string POReplace = "";
                        string POComments = woLine.POComments != null ? woLine.POComments : "";
                        woLine.POComments = rgxPOComm.Replace(POComments, POReplace);

                        woLine.ComponentGroupCode = !string.IsNullOrWhiteSpace(woLine.ComponentGroupCode) ? woLine.ComponentGroupCode : "-";
                        var own = owners.SingleOrDefault(x => x.Name.ToLower().Equals(woLine.ComponentGroupCode.ToLower()));
                        var ownz = own != null ? own.BOMShortageOwners.First().Name : woLine.ComponentGroupCode;
                        woLine.Owner = ownz;

                        decimal? demandUptoThisLine = new decimal(0.0);
                        decimal? stockLeftAfterThisDemand = new decimal(0.0);
                        decimal? so_demand = new decimal(0.0);

                        demandUptoThisLine = wolineTotals.Where(x => x.ComponentPart == woLine.ComponentPartNumber && x.PlannedIssueDate <= woLine.PlannedIssueDate).Sum(y => y.TotalDateDemand);
                        stockLeftAfterThisDemand = woLine.GoodStock - demandUptoThisLine;
                        so_demand = 0;

                        bool doesPODueDateMeetPID;
                        if (stockLeftAfterThisDemand < 0) // stock minus?
                        {
                            if (woLine.ReceiptDate <= woLine.PlannedIssueDate) // receipt date support planned issue date?
                            {
                                doesPODueDateMeetPID = stockLeftAfterThisDemand + woLine.QuantityPurchased >= 0; // po meet rest of demand? true/false
                            }
                            else
                            {
                                doesPODueDateMeetPID = false; // po date doesnt support
                            }
                        }
                        else
                        {
                            doesPODueDateMeetPID = true; // no problem
                        }

                        woLine.SO_Demand = so_demand;
                        woLine.DemandForThisDate = demandUptoThisLine;
                        woLine.StockLeftAfterThisDate = stockLeftAfterThisDemand;
                        woLine.POCoversDemand = doesPODueDateMeetPID;

                        // Process Cleaned Stock Locations

                        StringBuilder glocationBuilder = new StringBuilder();
                        StringBuilder blocationBuilder = new StringBuilder();
                        string glocation = string.Empty;
                        decimal gtotalQuantity = new decimal(0.0);
                        string blocation = string.Empty;
                        decimal btotalQuantity = new decimal(0.0);
                        //THAS_CONNECT_StockLocationCount_Result st1 = cleaned.SingleOrDefault(c => c.PartNumber.Equals(wo.ComponentPartNumber) && c.Location.Equals("STORE1"));
                        //THAS_CONNECT_StockLocationCount_Result st2 = cleaned.SingleOrDefault(c => c.PartNumber.Equals(wo.ComponentPartNumber) && c.Location.Equals("STORE2"));
                        //THAS_CONNECT_StockLocationCount_Result st3 = cleaned.SingleOrDefault(c => c.PartNumber.Equals(wo.ComponentPartNumber) && c.Location.Equals("STORE3"));
                        //THAS_CONNECT_StockLocationCount_Result good = cleaned.SingleOrDefault(cc => cc.PartNumber.Equals(wo.ComponentPartNumber) && !cc.Location.Equals("STORE1") && !cc.Location.Equals("STORE2") && !cc.Location.Equals("STORE3") && cc.isGood);
                        //THAS_CONNECT_StockLocationCount_Result bad = cleaned.SingleOrDefault(cc => cc.PartNumber.Equals(wo.ComponentPartNumber) && !cc.Location.Equals("STORE1") && !cc.Location.Equals("STORE2") && !cc.Location.Equals("STORE3") && !cc.isGood);
                        var ptgrp = stockCounts.Where(x => x.PartNumber.Equals(woLine.ComponentPartNumber));
                        if (ptgrp.Count() > 0)
                        {
                            ptgrp.ToList().ForEach(grp =>
                            {
                                if (grp.Location.Equals("STORE1"))
                                {
                                    woLine.Store1 = grp.On_Hand_Batch_Qty.Value;
                                }
                                else if (grp.Location.Equals("STORE2"))
                                {
                                    woLine.Store2 = grp.On_Hand_Batch_Qty.Value;
                                }
                                else if (grp.Location.Equals("STORE3"))
                                {
                                    woLine.Store3 = grp.On_Hand_Batch_Qty.Value;
                                }
                                else if (grp.Location.Equals("STORE4"))
                                {
                                    woLine.Store4= grp.On_Hand_Batch_Qty.Value;
                                }
                                else if (grp.Location.Equals("MOYFAB"))
                                {
                                    woLine.MoyFab = grp.On_Hand_Batch_Qty.Value;
                                }
                                else if (grp.Location.Equals("EAGLE OVERSEAS"))
                                {
                                    woLine.EagleOverseas = grp.On_Hand_Batch_Qty.Value;
                                }

                                else if (grp.Quarantined.Value == true || grp.ExcludeMRP.Value == true)
                                {
                                    blocationBuilder.Append(grp.On_Hand_Batch_Qty + " in " + grp.Location + " ");
                                    btotalQuantity += grp.On_Hand_Batch_Qty.Value;
                                }
                                else
                                {
                                    glocationBuilder.Append(grp.On_Hand_Batch_Qty + " in " + grp.Location + " ");
                                    gtotalQuantity += grp.On_Hand_Batch_Qty.Value;
                                }
                            }
                            );
                        }
                        glocation = glocationBuilder.ToString();
                        blocation = blocationBuilder.ToString();
                        //cleaned.Add(new THAS_CONNECT_StockLocationCount_Result { PartNumber = ptgrp.First().PartNumber, Location = glocation, On_Hand_Batch_Qty = gtotalQuantity, isGood = true });
                        //cleaned.Add(new THAS_CONNECT_StockLocationCount_Result { PartNumber = ptgrp.First().PartNumber, Location = blocation, On_Hand_Batch_Qty = btotalQuantity });
                        woLine.GoodLocations = glocation;
                        woLine.BadLocations = blocation;


                        ++ctr;
                        if (ctr == 1 || ctr % 5000 == 0)
                            Console.WriteLine("--- " + DateTime.Now + " sitting at " + ctr + " lines processed. ---");
                    }


                    // .Where(x => x.CurrentComponentMethodType == "Purchased")
                    resultSet.GroupBy(y => y.ComponentPartNumber).ToList().ForEach(pn =>
                    {
                        var list = pn.GroupBy(g => g.WOCommercialNotes).ToList();
                        list.ForEach(g =>
                        {
                            decimal curr_so_demand = g.Sum(x => x.Outstanding).Value;
                            WOLineReport wo = g.OrderBy(d => d.CompletionDate).Last();
                            ShortageLineGrouped export = new ShortageLineGrouped();
                            export.PartNo = wo.ComponentPartNumber;
                            export.Description = wo.ComponentPartDescription;
                            export.WorksOrderNumber = wo.WorksOrderNumber;
                            export.WODueDate = wo.CompletionDate;
                            export.ProductGroupCode = wo.ComponentGroupCode;
                            export.PartMethod = wo.CurrentComponentMethodType;
                            export.Responsibility = wo.WORespCode;
                            export.CommercialNotes = wo.WOCommercialNotes;
                            export.BatchNotes = wo.WOProductionNotes;
                            export.Demand = wo.Outstanding;
                            export.SO_Demand = curr_so_demand;
                            export.DemandForThisDate = wo.DemandForThisDate;
                            export.GoodStock = wo.GoodStock;
                            export.BadStock = wo.BadStock;
                            export.NetShortage = (export.GoodStock.Value - export.DemandForThisDate.Value);
                            export.StockLeftAfterThisDate = wo.StockLeftAfterThisDate;
                            export.Owner = wo.Owner;
                            export.Supplier = wo.SupplierName;
                            export.UnitCost = wo.UnitCost.Value;
                            export.PurchaseOrderNumber = wo.PurchaseOrderNumber;
                            export.PurchaseOrderDeliveryDate = wo.ReceiptDate;
                            export.PurchaseOrderQty = wo.QuantityPurchased;
                            export.ParentAssembly = wo.ParentPartNumber;
                            export.AllCallingWOs = g.Select(gg => gg.WorksOrderNumber).Aggregate((x, y) => x + ", " + y).ToString();
                            export.ParentAssemblyDescription = wo.ParentPartDescription;
                            export.Issued = wo.Issued;
                            export.POCoversDemand = wo.POCoversDemand;
                            export.Store1 = wo.Store1; //st1 != null ? st1.On_Hand_Batch_Qty.Value : new decimal(0.0);
                            export.Store2 = wo.Store2; //st2 != null ? st2.On_Hand_Batch_Qty.Value : new decimal(0.0);
                            export.Store3 = wo.Store3; //st3 != null ? st3.On_Hand_Batch_Qty.Value : new decimal(0.0);
                            export.Store4 = wo.Store4;
                            export.MoyFab = wo.MoyFab;
                            export.EagleOverseas = wo.EagleOverseas;
                            export.GoodLocations = wo.GoodLocations; //= good != null ? good.Location : "";
                            export.BadLocations = wo.BadLocations; //bad != null ? bad.Location : "";
                            export.POComments = wo.POComments;
                            export.compResponsibility = wo.CompRespCode;
                            export.IsStoresRequest = wo.IsStoresRequest.HasValue ? wo.IsStoresRequest.Value : false;
                            exports.Add(export);
                        });
                    });

                    exports = exports.OrderBy(pd => pd.PartNo).OrderBy(d => d.WODueDate).ToList();

                    var countz = 2;
                    foreach (var woLine in exports)
                    {
                        if ((woLine.GoodStock.Value - woLine.DemandForThisDate.Value) < new decimal(0.0))
                        {
                            partsws.Cells["R" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            partsws.Cells["R" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            partsws.Cells["R" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        }
                        else
                        {
                            partsws.Cells["R" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            partsws.Cells["R" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                            partsws.Cells["R" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        }
                        if (woLine.IsStoresRequest)
                        {
                            partsws.Cells["AL" + countz].Value = "Yes";
                        }
                        else
                        {
                            partsws.Cells["AL" + countz].Value = "No";
                        }
                        var realGoodStock = woLine.Store1 + woLine.Store2 + woLine.Store3 + woLine.Store4;
                        var offSiteStock = woLine.MoyFab + woLine.EagleOverseas;
                        var otherGoodLocStock = (woLine.GoodStock - realGoodStock - offSiteStock);

                        var isStoreEnough = realGoodStock >= woLine.DemandForThisDate;
                        var isStorePlusMoyfabEnough = realGoodStock + woLine.MoyFab >= woLine.DemandForThisDate;
                        var isStorePlusOffsiteEnough = realGoodStock + woLine.MoyFab + woLine.EagleOverseas >= woLine.DemandForThisDate;
                        var isOtherGoodLocationsEnough = otherGoodLocStock >= woLine.DemandForThisDate;
                        var isStorePlusOffsitePlusOtherGoodEnough = realGoodStock + offSiteStock + otherGoodLocStock >= woLine.DemandForThisDate;
                        var isAllGoodStockEnough = woLine.GoodStock >= woLine.DemandForThisDate;

                        //Check stock levels
                        var doesMoyfabHaveStock = woLine.MoyFab > 0 ? true : false;
                        var doesEagleOverseasHaveStock = woLine.EagleOverseas > 0 ? true : false;
                        var doesOtherGoodLocHaveStock = (woLine.GoodStock - realGoodStock - offSiteStock) > 0 ? true : false;

                        if (!isStoreEnough && (doesMoyfabHaveStock || doesEagleOverseasHaveStock || doesOtherGoodLocHaveStock))
                        {
                            if (!isStorePlusOffsitePlusOtherGoodEnough)
                            {
                                if (doesMoyfabHaveStock)
                                {
                                    partsws.Cells["AG" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AG" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AG" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                if (doesEagleOverseasHaveStock)
                                {
                                    partsws.Cells["AH" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AH" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AH" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                if (doesOtherGoodLocHaveStock)
                                {
                                    partsws.Cells["AI" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AI" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AI" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                            }
                            else if(isStorePlusOffsitePlusOtherGoodEnough)
                            {
                                if (isStorePlusMoyfabEnough && doesMoyfabHaveStock)
                                {
                                    partsws.Cells["AG" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AG" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AG" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (isStorePlusOffsiteEnough && doesEagleOverseasHaveStock && doesMoyfabHaveStock)
                                {
                                    partsws.Cells["AG" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AG" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AG" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    partsws.Cells["AH" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AH" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AH" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (isStorePlusOffsiteEnough && doesEagleOverseasHaveStock)
                                {
                                    partsws.Cells["AH" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AH" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AH" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (doesMoyfabHaveStock && doesEagleOverseasHaveStock && doesOtherGoodLocHaveStock)
                                {
                                    partsws.Cells["AG" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AG" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AG" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    partsws.Cells["AH" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AH" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AH" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    partsws.Cells["AI" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AI" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AI" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (doesMoyfabHaveStock && doesOtherGoodLocHaveStock)
                                {
                                    partsws.Cells["AG" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AG" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AG" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    partsws.Cells["AI" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AI" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AI" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (doesEagleOverseasHaveStock && doesOtherGoodLocHaveStock)
                                {
                                    partsws.Cells["AH" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AH" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AH" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    partsws.Cells["AI" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AI" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AI" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else if (doesOtherGoodLocHaveStock)
                                {
                                    partsws.Cells["AI" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    partsws.Cells["AI" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
                                    partsws.Cells["AI" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                            }
                        }

                        countz++;
                    }

                    partsws.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                    int rowCount = partsws.Dimension.Rows;
                    partsws.Column(2).Width = 30;
                    partsws.Cells["D2:D" + rowCount].Style.Numberformat.Format = "dd-mm-yyyy";
                    partsws.Cells["V2:V" + rowCount].Style.Numberformat.Format = "dd-mm-yyyy";
                    partsws.Cells["A1:AL1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    partsws.Cells["A1"].Value = "Part Number";
                    partsws.Cells["C1"].Value = "WO Number";
                    partsws.Cells["D1"].Value = "WO Due Date";
                    partsws.Cells["E1"].Value = "Product Group";
                    partsws.Cells["F1"].Value = "Part Method";
                    partsws.Cells["J1"].Value = "Commercial Notes";
                    partsws.Cells["K1"].Value = "Batch Notes";
                    partsws.Cells["O1"].Value = "Demand For This Date";
                    partsws.Cells["P1"].Value = "Good Stock";
                    partsws.Cells["Q1"].Value = "Bad Stock";
                    partsws.Cells["R1"].Value = "Net Shortage";
                    partsws.Cells["S1"].Value = "Stock After This Date";
                    partsws.Cells["T1"].Value = "All Calling WOs";
                    partsws.Cells["U1"].Value = "PO Number";
                    partsws.Cells["V1"].Value = "PO Acknowledge Date";
                    partsws.Cells["W1"].Value = "PO Quantity";
                    partsws.Cells["X1"].Value = "PO Comments";

                    partsws.Cells["Y1"].Value = "Parent Assembly";
                    partsws.Cells["Z1"].Value = "Parent Assembly Description";
                    partsws.Cells["AB1"].Value = "PO Covers Demand?";

                    partsws.Cells["AC1"].Value = "Store 1";
                    partsws.Cells["AD1"].Value = "Store 2";
                    partsws.Cells["AE1"].Value = "Store 3";
                    partsws.Cells["AF1"].Value = "Store 4";
                    partsws.Cells["AG1"].Value = "MoyFab";
                    partsws.Cells["AH1"].Value = "Eagle Overseas";
                    partsws.Cells["AI1"].Value = "Other Good Locations";
                    partsws.Cells["AJ1"].Value = "Bad Locations";
                    partsws.Cells["AK1"].Value = "Comp Resp";
                    partsws.Cells["AL1"].Value = "Stores Request?";

                    partsws.Cells["A1:AL1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DodgerBlue);
                    partsws.Cells["M1:O1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                    partsws.Cells["P1:Q1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                    partsws.Cells["R1:R1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                    partsws.Cells["AC:AJ1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);

                    partsws.Cells["R2:R"+ rowCount].Style.Numberformat.Format = "0.00";

                    //partsws.Cells[partsws.Dimension.Address].AutoFitColumns();
                    partsws.DefaultColWidth = 15.0;
                    partsws.Column(1).Width = 30.0;
                    partsws.Column(2).Width = 30.0;
                    partsws.Column(3).Width = 15.0;
                    partsws.Column(4).Width = 15.0;
                    partsws.Column(5).Width = 20.0;
                    partsws.Column(6).Width = 20.0;
                    partsws.Column(7).Width = 20.0;
                    partsws.Column(8).Width = 20.0;
                    partsws.Column(9).Width = 20.0;
                    partsws.Column(10).Width = 30.0;
                    partsws.Column(11).Width = 30.0;
                    partsws.Column(12).Width = 10.0;
                    partsws.Column(13).Width = 15.0;
                    partsws.Column(14).Width = 15.0;
                    partsws.Column(15).Width = 20.0;
                    partsws.Column(16).Width = 12.5;
                    partsws.Column(17).Width = 12.5;
                    partsws.Column(18).Width = 15.0;
                    partsws.Column(20).Width = 30.0;
                    partsws.Column(24).Width = 30.0;
                    partsws.Column(25).Width = 30.0;
                    partsws.Column(32).Width = 20.0;
                    partsws.Column(33).Width = 20.0;
                    partsws.Column(34).Width = 20.0;
                    partsws.Column(35).Width = 50.0;
                    partsws.Column(36).Width = 50.0;
                    partsws.View.ZoomScale = 75;
                    partsws.DeleteColumn(12);

                    // Generate the WO-Parts-Level worksheet report.
                    var workSheet = excelPackage.Workbook.Worksheets.Add("Shorts-WO-Parts-Level");

                    resultSet.ToList().ForEach(wo =>
                    {
                        ShortageLine export = new ShortageLine();
                        export.PartNo = wo.ComponentPartNumber;
                        export.Description = wo.ComponentPartDescription;
                        export.WorksOrderNumber = wo.WorksOrderNumber;
                        export.WODueDate = wo.CompletionDate;
                        export.ProductGroupCode = wo.ComponentGroupCode;
                        export.PartMethod = wo.CurrentComponentMethodType;
                        export.Responsibility = wo.WORespCode;
                        export.CommercialNotes = wo.WOCommercialNotes;
                        export.BatchNotes = wo.WOProductionNotes;
                        export.Demand = wo.Outstanding;
                        export.DemandForThisDate = wo.DemandForThisDate;
                        export.GoodStock = wo.GoodStock;
                        export.BadStock = wo.BadStock;
                        export.NetShortage = (export.GoodStock.Value - export.DemandForThisDate.Value);
                        export.StockLeftAfterThisDate = wo.StockLeftAfterThisDate;
                        export.Supplier = wo.SupplierName;
                        export.PurchaseOrderNumber = wo.PurchaseOrderNumber;
                        export.PurchaseOrderDeliveryDate = wo.ReceiptDate;
                        export.PurchaseOrderQty = wo.QuantityPurchased;
                        export.PORaisedBy = wo.PORaisedBy;
                        export.ComponentWorksOrder = wo.ComponentWorksOrder;
                        export.ComponentWODueDate = wo.WODueDate;
                        export.ComponentWOQuantity = wo.Quantity;
                        export.WORaisedBy = wo.WORaisedBy;
                        export.ParentAssembly = wo.ParentPartNumber;
                        export.ParentAssemblyDescription = wo.ParentPartDescription;
                        export.Issued = wo.Issued;
                        export.POCoversDemand = wo.POCoversDemand;
                        export.UnitCost = wo.UnitCost.Value;
                        export.Store1 = wo.Store1; //st1 != null ? st1.On_Hand_Batch_Qty.Value : new decimal(0.0);
                        export.Store2 = wo.Store2; //st2 != null ? st2.On_Hand_Batch_Qty.Value : new decimal(0.0);
                        export.Store3 = wo.Store3; //st3 != null ? st3.On_Hand_Batch_Qty.Value : new decimal(0.0);
                        export.Store4 = wo.Store4;
                        export.MoyFab = wo.MoyFab;
                        export.EagleOverseas = wo.EagleOverseas;
                        export.GoodLocations = wo.GoodLocations; //= good != null ? good.Location : "";
                        export.BadLocations = wo.BadLocations; //bad != null ? bad.Location : "";
                        export.compResponsibility = wo.CompRespCode;
                        export.IsStoresRequest = wo.IsStoresRequest.HasValue ? wo.IsStoresRequest.Value : false;
                        exports2.Add(export);
                    });

                    exports2 = exports2.OrderBy(pd => pd.PartNo).OrderBy(d => d.WODueDate).ToList();
                    countz = 2;

                    foreach (var woLine in exports2)
                    {
                        if (woLine.IsStoresRequest)
                        {
                            workSheet.Cells["AL" + countz].Value = "Yes";
                        }
                        else
                        {
                            workSheet.Cells["AL" + countz].Value = "No";
                        }

                        if ((woLine.GoodStock.Value - woLine.DemandForThisDate.Value) < new decimal(0.0))
                        {
                            workSheet.Cells["P" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells["P" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            workSheet.Cells["P" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        }
                        else
                        {
                            workSheet.Cells["P" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells["P" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                            workSheet.Cells["P" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        }
                        countz++;
                    }

                    workSheet.Cells["A1"].LoadFromCollection(exports2, true, OfficeOpenXml.Table.TableStyles.Medium2);
                    rowCount = workSheet.Dimension.Rows;
                    workSheet.Cells["D2:D" + rowCount].Style.Numberformat.Format = "dd-mm-yyyy";
                    workSheet.Cells["S2:S" + rowCount].Style.Numberformat.Format = "dd-mm-yyyy";
                    workSheet.Cells["W2:W" + rowCount].Style.Numberformat.Format = "dd-mm-yyyy";
                    workSheet.Cells["A1:AM1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells["A1"].Value = "Part Number";
                    workSheet.Cells["C1"].Value = "WO Number";
                    workSheet.Cells["D1"].Value = "WO Due Date";
                    workSheet.Cells["E1"].Value = "WO Raised By";
                    workSheet.Cells["F1"].Value = "Product Group";
                    workSheet.Cells["G1"].Value = "Part Method";
                    workSheet.Cells["J1"].Value = "Commercial Notes";
                    workSheet.Cells["K1"].Value = "Batch Notes";
                    workSheet.Cells["I1"].Value = "Owner/Supplier";
                    workSheet.Cells["M1"].Value = "Demand For This Date";
                    workSheet.Cells["N1"].Value = "Good Stock";
                    workSheet.Cells["O1"].Value = "Bad Stock";
                    workSheet.Cells["P1"].Value = "Net Shortage";
                    workSheet.Cells["Q1"].Value = "Stock After This Date";
                    workSheet.Cells["R1"].Value = "PO Number";
                    workSheet.Cells["S1"].Value = "PO Acknowledge Date";
                    workSheet.Cells["T1"].Value = "PO Quantity";
                    workSheet.Cells["U1"].Value = "PO Raised By";
                    workSheet.Cells["V1"].Value = "Component WO";
                    workSheet.Cells["W1"].Value = "Component WO Due Date";
                    workSheet.Cells["X1"].Value = "Component WO Quantity";
                    workSheet.Cells["Y1"].Value = "Parent Assembly";
                    workSheet.Cells["Z1"].Value = "Parent Assembly Description";
                    workSheet.Cells["AB1"].Value = "PO Covers Demand?";
                    workSheet.Cells["AC1"].Value = "Unit Cost";
                    workSheet.Cells["AD1"].Value = "Store 1";
                    workSheet.Cells["AE1"].Value = "Store 2";
                    workSheet.Cells["AF1"].Value = "Store 3";
                    workSheet.Cells["AG1"].Value = "Store 4";
                    workSheet.Cells["AH1"].Value = "MoyFab";
                    workSheet.Cells["AI1"].Value = "Eagle Overseas";
                    workSheet.Cells["AJ1"].Value = "Other Good Locations";
                    workSheet.Cells["AK1"].Value = "Bad Locations";
                    workSheet.Cells["AL1"].Value = "Comp Resp";
                    workSheet.Cells["AM1"].Value = "Stores Request?";

                    workSheet.Cells["A1:AM1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DodgerBlue);
                    workSheet.Cells["L1:M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                    workSheet.Cells["N:O1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                    workSheet.Cells["P1:P1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                    workSheet.Cells["R1:U1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                    workSheet.Cells["V1:X1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                    workSheet.Cells["R1:X1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);

                    workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                    workSheet.Column(1).Width = 25.0;
                    workSheet.Column(2).Width = 40.0;
                    workSheet.Column(9).Width = 25.0;
                    workSheet.Column(10).Width = 30.0;
                    workSheet.Column(11).Width = 30.0;
                    workSheet.View.ZoomScale = 75;
                    workSheet.DeleteColumn(29);

                    //var stockSheet = excelPackage.Workbook.Worksheets.Add("Parts-Stock-Locations");
                    //stockSheet.Cells["A1"].LoadFromCollection(cleaned, true, OfficeOpenXml.Table.TableStyles.Medium2);
                    //rowCount = stockSheet.Dimension.Rows;
                    //stockSheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //stockSheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DodgerBlue);
                    //stockSheet.Cells["A1"].Value = "Part Number";
                    //stockSheet.Cells["B1"].Value = "Total On Hand Qty";
                    //stockSheet.Cells["C1"].Value = "Location Details";
                    ////stockSheet.Cells["C1:C"].IsRichText = true;
                    //stockSheet.Column(1).Width = 30.0;
                    //stockSheet.Column(2).Width = 20.0;
                    //stockSheet.Column(3).Width = 250.0;
                    //stockSheet.View.ZoomScale = 75;

                    excelPackage.Save();
                }
                OverwriteGenericCopy(fileInfo.Name, theDate);

                connect.Database.ExecuteSqlCommand("truncate table WOLineReport_WOPartsLevel");
                connect.Database.ExecuteSqlCommand("truncate table WOLineReport_PartsShortages");
                CopySecondSheetToDB(exports2);
                CopyFirstSheetToDB(exports);
            }
        }

        //private static bool CreateDirectoryStructure(out FileInfo fileInfo, string date, string dateHours)
        //{
        //    fileInfo = new FileInfo(string.Format(@"\\THAS-NAS01\Shortages$\_New Shortage Reports\{0}\InBuildShortageReport_{1}.xlsx", date, dateHours));
        //    //fileInfo = new FileInfo(string.Format(@"S:\Shortages\_New Shortage Reports\Test\InBuildShortageReport_{1}.xlsx", date, dateHours));
        //    try
        //    {
        //        var fullpath = string.Format(@"\\THAS-NAS01\Shortages$\_New Shortage Reports\{0}\InBuildShortageReport_{1}.xlsx", date, dateHours);
        //        //var fullpath = string.Format(@"S:\Shortages\_New Shortage Reports\Test\InBuildShortageReport_{1}.xlsx", date, dateHours);
        //        if (!File.Exists(fullpath))
        //        {
        //            fileInfo = new FileInfo(fullpath);
        //            fileInfo.Directory.Create();
        //            return true;
        //        }
        //        else
        //            return false; // get out of here.              
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Issue : " + ex.Message);
        //        return false;
        //    }
        //}

        private static void CopyFirstSheetToDB(List<ShortageLineGrouped> dataSet)
        {
            ConnectReportDbEntities connect = null;
            try
            {
                connect = new ConnectReportDbEntities();
                connect.Configuration.AutoDetectChangesEnabled = false;

                int count = 0;
                foreach (var line in dataSet)
                {
                    ++count;
                    var gp = new WOLineReport_PartsShortages();
                    gp.PartNumber = line.PartNo;
                    gp.Description = line.Description;
                    gp.WONumber = line.WorksOrderNumber;
                    gp.WODueDate = line.WODueDate;
                    gp.ProductGroup = line.ProductGroupCode;
                    gp.PartMethod = line.PartMethod;
                    gp.Responsibility = line.Responsibility;
                    gp.CommercialNotes = line.CommercialNotes;
                    gp.BatchNotes = line.BatchNotes;
                    gp.Demand = line.Demand;
                    gp.SODemand = line.SO_Demand;
                    gp.DemandForThisDate = line.DemandForThisDate;
                    gp.GoodStock = line.GoodStock;
                    gp.BadStock = line.BadStock;
                    gp.NetShortage = line.NetShortage;
                    gp.StockAfterThisDate = line.StockLeftAfterThisDate;
                    gp.Supplier = line.Supplier;
                    gp.PONumber = line.PurchaseOrderNumber;
                    gp.PODeliveryDate = line.PurchaseOrderDeliveryDate;
                    gp.POQuantity = line.PurchaseOrderQty;
                    gp.ParentAssembly = line.ParentAssembly;
                    gp.ParentAssemblyDescription = line.ParentAssemblyDescription;
                    gp.Issued = line.Issued;
                    gp.POCoversDemand = line.POCoversDemand;
                    gp.Owner = line.Owner;
                    gp.UnitCost = line.UnitCost;
                    gp.Store1 = line.Store1;
                    gp.Store2 = line.Store2;
                    gp.Store3 = line.Store3;
                    gp.OtherGood = line.GoodLocations;
                    gp.OtherBad = line.BadLocations;
                    gp.AllWOs = line.AllCallingWOs;
                    gp.CompRespCode = line.compResponsibility;
                    gp.IsStoresRequest = line.IsStoresRequest;
                    connect = AddToContextFirst(connect, gp, count, 500, true);
                }
                connect.SaveChanges();
            }
            finally
            {
                if (connect != null)
                    connect.Dispose();
            }
        }

        private static void CopySecondSheetToDB(List<ShortageLine> dataSet)
        {
            ConnectReportDbEntities connect = null;
            try
            {
                connect = new ConnectReportDbEntities();
                connect.Configuration.AutoDetectChangesEnabled = false;

                int count = 0;
                foreach (var line in dataSet)
                {
                    ++count;
                    var gp = new WOLineReport_WOPartsLevel();
                    gp.PartNumber = line.PartNo;
                    gp.Description = line.Description;
                    gp.WONumber = line.WorksOrderNumber;
                    gp.WODueDate = line.WODueDate;
                    gp.ProductGroup = line.ProductGroupCode;
                    gp.PartMethod = line.PartMethod;
                    gp.Responsibility = line.Responsibility;
                    gp.CommercialNotes = line.CommercialNotes;
                    gp.BatchNotes = line.BatchNotes;
                    gp.Demand = line.Demand;
                    gp.DemandForThisDate = line.DemandForThisDate;
                    gp.GoodStock = line.GoodStock;
                    gp.BadStock = line.BadStock;
                    gp.NetShortage = line.NetShortage;
                    gp.StockAfterThisDate = line.StockLeftAfterThisDate;
                    gp.Supplier = line.Supplier;
                    gp.PONumber = line.PurchaseOrderNumber;
                    gp.PODeliveryDate = line.PurchaseOrderDeliveryDate;
                    gp.POQuantity = line.PurchaseOrderQty;
                    gp.PORaisedBy = line.PORaisedBy;
                    gp.ComponentWO = line.ComponentWorksOrder;
                    gp.ComponentWODueDate = line.ComponentWODueDate;
                    gp.ComponentWOQuantity = line.ComponentWOQuantity;
                    gp.WORaisedBy = line.WORaisedBy;
                    gp.ParentAssembly = line.ParentAssembly;
                    gp.ParentAssemblyDescription = line.ParentAssemblyDescription;
                    gp.Issued = line.Issued;
                    gp.POCoversDemand = line.POCoversDemand;
                    gp.Owner = string.Empty;
                    gp.UnitCost = line.UnitCost;
                    gp.Store1 = line.Store1;
                    gp.Store2 = line.Store2;
                    gp.Store3 = line.Store3;
                    gp.OtherGood = line.GoodLocations;
                    gp.OtherBad = line.BadLocations;
                    gp.AllWOs = string.Empty;
                    gp.CompRespCode = line.compResponsibility;
                    gp.IsStoresRequest = line.IsStoresRequest;
                    connect = AddToContextSecond(connect, gp, count, 500, true);
                }
                connect.SaveChanges();
            }
            finally
            {
                if (connect != null)
                    connect.Dispose();
            }
        }

        private static bool OverwriteGenericCopy(string newlyGeneratedFilename, string date)
        {
            try
            {

                var directory = @"\\tas\reports$\Shortage Reports\Without Costing Info\Generic\";
                Directory.CreateDirectory(directory);
                var filename = "InBuildShortageReport.xlsx";
                FileInfo checkFile = new FileInfo(directory + filename);

                if (checkFile.Exists)
                {
                    try
                    {
                        checkFile.IsReadOnly = false;
                        File.Delete(directory + filename);
                    }
                    catch (Exception ex)
                    {
                        return false;
                    }
                }
                var fileInfo = new FileInfo(string.Format(@"\\tas\reports$\Shortage Reports\Without Costing Info\{0}\{1}", date, newlyGeneratedFilename));
                fileInfo.CopyTo(directory + filename);

                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool CreateDirectoryStructure(out FileInfo fileInfo, string date, string dateHours, string filename, string folderPath, bool costed)
        {
            string path = @"\\tas\reports$\{0}\{1}\";
            if (costed)
            {
                path = @"\\tas\reports$\{0}\With Costing Info\{1}\";
            }
            else
            {
                path = @"\\tas\reports$\{0}\Without Costing Info\{1}\";
            }


            fileInfo = new FileInfo(string.Format(path + filename + "_{2}.xlsx", folderPath, date, dateHours));
            try
            {
                var fullpath = string.Format(path + filename + "_{2}.xlsx", folderPath, date, dateHours);
                if (!File.Exists(fullpath))
                {
                    fileInfo = new FileInfo(fullpath);
                    fileInfo.Directory.Create();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue : " + ex.Message);
                return false;
            }
        }


        private static ConnectReportDbEntities AddToContextFirst(ConnectReportDbEntities context, WOLineReport_PartsShortages entity, int count, int commitCount, bool recreateContext)
        {
            context.Set<WOLineReport_PartsShortages>().Add(entity);

            if (count % commitCount == 0)
            {
                context.SaveChanges();
                if (recreateContext)
                {
                    context.Dispose();
                    context = new ConnectReportDbEntities();
                    context.Configuration.AutoDetectChangesEnabled = false;
                }
            }
            return context;
        }

        private static ConnectReportDbEntities AddToContextSecond(ConnectReportDbEntities context, WOLineReport_WOPartsLevel entity, int count, int commitCount, bool recreateContext)
        {
            context.Set<WOLineReport_WOPartsLevel>().Add(entity);

            if (count % commitCount == 0)
            {
                context.SaveChanges();
                if (recreateContext)
                {
                    context.Dispose();
                    context = new ConnectReportDbEntities();
                    context.Configuration.AutoDetectChangesEnabled = false;
                }
            }
            return context;
        }

        //static void sendMail(string errorMessage)
        //{
        //    MailMessage mail = new MailMessage("OpenWOReportV2@thompsonaero.com", "sean.kelly@thompsonaero.com");
        //    SmtpClient client = new SmtpClient();
        //    client.Port = 25;
        //    client.Host = "remote.thompsonaero.com";
        //    mail.Subject = "OpenWOReportV2 Export Has Failed.";
        //    mail.Body = "Failed For The Following Reason: " + errorMessage;
        //    client.Send(mail);
        //}
    }
}
