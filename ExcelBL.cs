using Common.Constants;
using CsvHelper;
using DataAccessLayer.DBModel;
using ExcelFormatting.Model;
//using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFormatting
{
    public class ExcelBL
    {
        //private static Logger logger = LogManager.GetCurrentClassLogger();
        string folderPath = string.Empty;
        public ExcelBL()
        {
            this.folderPath = ConfigurationManager.AppSettings["ExcelFolder"];
        }

        //public string RCPFromSales { get { return ConfigurationManager.AppSettings["RCPFromSales"]; } }
        //public string FormattedRCPFromSales { get { return ConfigurationManager.AppSettings["FormattedRCPFromSales"]; } }
        //public string Salesman { get { return ConfigurationManager.AppSettings["Salesman"]; } }
        //public string Beat { get { return ConfigurationManager.AppSettings["Beat"]; } }
        //public string SalesmanRoute { get { return ConfigurationManager.AppSettings["SalesmanRoute"]; } }
        //public string Outlet { get { return ConfigurationManager.AppSettings["Outlet"]; } }
        //public string CPCategory { get { return ConfigurationManager.AppSettings["CPCategory"]; } }
        //public string CustomerRoute { get { return ConfigurationManager.AppSettings["CustomerRoute"]; } }
        //public string BeatPanning { get { return ConfigurationManager.AppSettings["BeatPanning"]; } }

        /// <summary>
        /// Method to check the excel if it is older or newer
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool IsOlderVersionExcel(string fileName)
        {
            bool result = false;
            try
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(this.folderPath);
                foreach (FileInfo file in di.GetFiles(fileName + ".xls"))
                {
                    result = file != null ? true : false;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return result;
        }

        private bool IsOlderVersionExcelNew(string fileName)
        {
            bool result = false;
            try
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(this.folderPath);
                foreach (FileInfo file in di.GetFiles(fileName))
                {
                    result = file.FullName.EndsWith(".xls") ? true : false;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return result;
        }

        /// <summary>
        /// Method to read data from DMS file
        /// </summary>
        /// <returns></returns>
        public List<RCPFromDMSModel> ReadRCPFromDMS(long processInstanceId, out string errorMessage)
        {
            errorMessage = string.Empty;
            string fileName = string.Empty, sheetName = string.Empty;//"RCPReport21022019085648.csv",
            GetFileName(processInstanceId, ref fileName, ref sheetName, Common.Constants.Process.States.RCPExccelProcess.DMSAttachmentDownload);
            List<RCPFromDMSModel> rcpFromDMSModelList = new List<RCPFromDMSModel>();
            DataTable data = new DataTable();
            try
            {
                using (StreamReader reader = new StreamReader(folderPath + fileName))
                {
                    string[] headerRow = new string[] { };
                    using (var csv = new CsvReader(reader))
                    {
                        csv.Read();
                        csv.ReadHeader();
                        headerRow = csv.Context.HeaderRecord;
                        for (int loop = 0; loop < csv.Context.HeaderRecord.Count(); loop++)
                        {
                            csv.Context.HeaderRecord[loop] = Regex.Replace(csv.Context.HeaderRecord[loop], @"\s", string.Empty);
                        }
                        csv.ReadHeader();
                        csv.Configuration.Delimiter = ",";
                        var records = csv.GetRecords<RCPFromDMSModel>().ToList();
                        rcpFromDMSModelList = records.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            if (rcpFromDMSModelList.Count <= 0)
                errorMessage = "{\"Detail\":{\"" + Common.Constants.JSON.Tags.Message.Details.Key + "\":\"" + "RCP file is not exists or is empty." + "\"}}";
            return rcpFromDMSModelList;
        }

        /// <summary>
        /// Method to read data from sales team file
        /// </summary>
        /// <returns></returns>
        public List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam(long processInstanceId, out string errorMessage)
        {
            errorMessage = string.Empty;
            List<RCPFromSalesTeamModel> rcpFromSalesTeamModelList = new List<RCPFromSalesTeamModel>();
            DataTable data = null;
            string fileName = string.Empty, sheetName = string.Empty, excelFileName = string.Empty, excelFilePath = string.Empty;
            GetFileName(processInstanceId, ref fileName, ref sheetName, Common.Constants.Process.States.RCPExccelProcess.MailAttachmentDownload);
            try
            {
                bool isOlderExcel = IsOlderVersionExcel(fileName.Replace(".xlsx", ""));
                if (isOlderExcel)
                {
                    var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "FromSalesTeam");
                    data = ds.Tables["FromSalesTeam"];
                    rcpFromSalesTeamModelList = ConvertDataTable<RCPFromSalesTeamModel>(data);
                }
                else
                {
                    fileName = this.folderPath + fileName;
                    FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
                    DataTable dTable = new DataTable();
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        //ExcelWorksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
                        if (workSheet != null && workSheet.Dimension != null)
                        {
                            int totalRows = workSheet.Dimension.Rows;
                            for (int i = 2; i <= totalRows; i++)
                            {
                                rcpFromSalesTeamModelList.Add(new RCPFromSalesTeamModel
                                {
                                    RegionName = workSheet.Cells[i, 1].Value != null ? workSheet.Cells[i, 1].Value.ToString() : string.Empty,
                                    State = workSheet.Cells[i, 2].Value != null ? workSheet.Cells[i, 2].Value.ToString() : string.Empty,
                                    City = workSheet.Cells[i, 3].Value != null ? workSheet.Cells[i, 3].Value.ToString() : string.Empty,
                                    TownName = workSheet.Cells[i, 4].Value != null ? workSheet.Cells[i, 4].Value.ToString() : string.Empty,
                                    ReportingTo = workSheet.Cells[i, 5].Value != null ? workSheet.Cells[i, 5].Value.ToString() : string.Empty,
                                    SalesForceCode = workSheet.Cells[i, 6].Value != null ? Convert.ToInt32(workSheet.Cells[i, 6].Value) : 0,
                                    DistributedBranchCode = workSheet.Cells[i, 7].Value != null ? workSheet.Cells[i, 7].Value.ToString() : string.Empty,
                                    RouteName = workSheet.Cells[i, 8].Value != null ? workSheet.Cells[i, 8].Value.ToString() : string.Empty,
                                    IsNewRoute = workSheet.Cells[i, 9].Value != null ? workSheet.Cells[i, 9].Value.ToString() : string.Empty,
                                    RouteCode = workSheet.Cells[i, 10].Value != null ? workSheet.Cells[i, 10].Value.ToString() : string.Empty,
                                    SEType = workSheet.Cells[i, 11].Value != null ? workSheet.Cells[i, 11].Value.ToString() : string.Empty,
                                    SECode = workSheet.Cells[i, 12].Value != null ? workSheet.Cells[i, 12].Value.ToString() : string.Empty,
                                    IsNewSalesman = workSheet.Cells[i, 13].Value != null ? workSheet.Cells[i, 13].Value.ToString() : string.Empty,
                                    SalesmanCategory = workSheet.Cells[i, 14].Value != null ? workSheet.Cells[i, 14].Value.ToString() : string.Empty,
                                    DayOfWeek = workSheet.Cells[i, 15].Value != null ? workSheet.Cells[i, 15].Value.ToString() : string.Empty,
                                    OutletCode = workSheet.Cells[i, 16].Value != null ? workSheet.Cells[i, 16].Value.ToString() : string.Empty,
                                    OutletName = workSheet.Cells[i, 17].Value != null ? workSheet.Cells[i, 17].Value.ToString() : string.Empty,
                                    OutletAddress = workSheet.Cells[i, 18].Value != null ? workSheet.Cells[i, 18].Value.ToString() : string.Empty,
                                    ProductHierarchyCategoryName = workSheet.Cells[i, 19].Value != null ? workSheet.Cells[i, 19].Value.ToString() : string.Empty,
                                    PostalCode = workSheet.Cells[i, 20].Value != null ? Convert.ToInt32(workSheet.Cells[i, 20].Value) : 123456,
                                    Retlrtype = workSheet.Cells[i, 21].Value != null ? workSheet.Cells[i, 21].Value.ToString() : string.Empty,
                                    CustChannelType = workSheet.Cells[i, 22].Value != null ? workSheet.Cells[i, 22].Value.ToString() : string.Empty,
                                    CustChannelSubType = workSheet.Cells[i, 23].Value != null ? workSheet.Cells[i, 23].Value.ToString() : string.Empty,
                                    OutletPhoneNo = workSheet.Cells[i, 24].Value != null ? Convert.ToInt64(workSheet.Cells[i, 24].Value) : 1234567890,
                                    StoreType = workSheet.Cells[i, 25].Value != null ? workSheet.Cells[i, 25].Value.ToString() : string.Empty,
                                    OutletIdForRemoval = workSheet.Cells[i, 26].Value != null ? workSheet.Cells[i, 26].Value.ToString() : string.Empty,
                                    RouteCodeforOutletIdRemoval = workSheet.Cells[i, 27].Value != null ? workSheet.Cells[i, 27].Value.ToString() : string.Empty,
                                    OutletIdforRouteTransfer = workSheet.Cells[i, 28].Value != null ? workSheet.Cells[i, 28].Value.ToString() : string.Empty,
                                    RouteCodeforOutletRouteTransfer = workSheet.Cells[i, 29].Value != null ? workSheet.Cells[i, 29].Value.ToString() : string.Empty,
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            if (rcpFromSalesTeamModelList.Count <= 0)
                errorMessage = "{\"Detail\":{\"" + Common.Constants.JSON.Tags.Message.Details.Key + "\":\"" + "File from sales team is not exists or is empty." + "\"}}";
            return rcpFromSalesTeamModelList;
        }

        public void MoveFileToFolder(string sourcePath, string targetPath)
        {
            try
            {
                if (Directory.Exists(sourcePath))
                {
                    foreach (var file in new DirectoryInfo(sourcePath).GetFiles())
                    {
                        file.MoveTo($@"{targetPath}\{file.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName].ToString(), null);
                    else
                        continue;
                }
            }
            return obj;
        }

        /// <summary>
        /// Method to apply rules
        /// </summary>
        /// <param name="ReadRCPFromDMS"></param>
        /// <param name="ReadRCPFromSalesTeam"></param>
        public string ApplyingRules(List<RCPFromDMSModel> ReadRCPFromDMS, ref List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, ref List<RCPFromSalesTeamModel> lstFromSalesNewOutlets, ref List<RCPFromSalesTeamModel> outletRouteRemovalList, ref List<RCPFromSalesTeamModel> lstOutletRouteTransfer, out string errorMessage)
        {
            errorMessage = string.Empty;
            string branchCode = string.Empty;
            string fileRequiredJSON = "{ ";
            try
            {
                using (GPILEntities entities = new GPILEntities())
                {
                    RCPFromSalesTeamModel tempRcpFromSalesTeamModel = new RCPFromSalesTeamModel();

                    outletRouteRemovalList = ReadRCPFromSalesTeam.Where(x => !string.IsNullOrWhiteSpace(x.OutletIdForRemoval) && !string.IsNullOrWhiteSpace(x.RouteCodeforOutletIdRemoval)).ToList();
                    lstOutletRouteTransfer=ReadRCPFromSalesTeam.Where(x => !string.IsNullOrWhiteSpace(x.OutletIdforRouteTransfer) && !string.IsNullOrWhiteSpace(x.RouteCodeforOutletRouteTransfer)).ToList();
                    lstOutletRouteTransfer = ReadRCPFromSalesTeam.Where(x => x.OutletIdForRemoval == x.OutletIdforRouteTransfer && x.RouteCodeforOutletIdRemoval != x.RouteCodeforOutletRouteTransfer).ToList();

                    ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(s => !string.IsNullOrWhiteSpace(s.DayOfWeek)
                    && !string.IsNullOrWhiteSpace(s.DistributedBranchCode) && !string.IsNullOrWhiteSpace(s.RegionName)
                    && !string.IsNullOrWhiteSpace(s.RouteName) && !string.IsNullOrWhiteSpace(s.SECode)
                    && !string.IsNullOrWhiteSpace(s.State) && !string.IsNullOrWhiteSpace(s.City)).ToList();

                    ReadRCPFromSalesTeam.RemoveAll(x => x.IsNewSalesman.ToLower().Equals("new") && (string.IsNullOrWhiteSpace(x.SalesForceCode.ToString())
                    || string.IsNullOrWhiteSpace(x.ReportingTo) || string.IsNullOrWhiteSpace(x.SEType) || string.IsNullOrWhiteSpace(x.SalesmanCategory)));

                    ReadRCPFromSalesTeam.RemoveAll(x => !(x.IsNewRoute.ToLower().Equals("new")) && string.IsNullOrWhiteSpace(x.RouteCode));

                    ReadRCPFromSalesTeam.RemoveAll(x => x.OutletCode.ToLower().Equals("new") && ((string.IsNullOrWhiteSpace(x.OutletName)
                    || string.IsNullOrWhiteSpace(x.OutletAddress) || string.IsNullOrWhiteSpace(x.CustChannelType) || string.IsNullOrWhiteSpace(x.CustChannelSubType)
                    || string.IsNullOrWhiteSpace(x.StoreType))));

                    //ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(p => ReadRCPFromDMS.Any(x => x.DistrCode == p.DistributedBranchCode)).ToList();

                    if (ReadRCPFromSalesTeam.Count > 0)
                    {
                        lstFromSalesNewOutlets = ReadRCPFromSalesTeam.Where(x => x.OutletCode.ToLower().Contains("new") || string.IsNullOrEmpty(x.OutletCode)).ToList();
                        List<RCPFromSalesTeamModel> NewSECodelst = ReadRCPFromSalesTeam.Where(x => x.IsNewSalesman.ToLower().Contains("new")).ToList();
                        List<RCPFromSalesTeamModel> NewRouteCodeslst = ReadRCPFromSalesTeam.Where(x => x.IsNewRoute.ToLower().Contains("new")).OrderBy(x => x.RouteName).ToList();
                        List<string> excelRules = entities.TBL_ExcelRules.Where(x => x.IsActive != null && x.IsActive == true).Select(x => x.RuleName).ToList();

                        if (excelRules != null && excelRules.Contains("SECode"))
                        {
                            if (NewSECodelst != null && NewSECodelst.Count > 0)
                            {
                                fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.Salesman.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareSalesmanExcel + "\",";
                            }
                        }
                        if (excelRules != null && excelRules.Contains("CreateRouteCode"))
                        {
                            if (NewRouteCodeslst != null && NewRouteCodeslst.Count() > 0)
                            {
                                int count = 1; string RouteCode = string.Empty, RouteName = string.Empty;
                                foreach (RCPFromSalesTeamModel item in NewRouteCodeslst)
                                {
                                    if (string.IsNullOrWhiteSpace(RouteName) || !(item.RouteName.Equals(RouteName)))
                                    {
                                        RouteName = item.RouteName;
                                        branchCode = item.RegionName.Trim();
                                        string bCode = GetBranchCode(branchCode);//Enum.Parse(typeof(BranchCode), branchCode).ToString();
                                        RouteCode = bCode + DateTime.Now.AddYears(1).ToString("ddMMyy") + count.ToString().PadLeft(3, '0');
                                        count++;
                                        //item.RouteCode = RouteCode;
                                    }
                                    item.RouteCode = RouteCode;

                                    var index = ReadRCPFromSalesTeam.IndexOf(item);
                                    if (index != -1)
                                    {
                                        ReadRCPFromSalesTeam[index].RouteCode = RouteCode;
                                    }
                                }
                                fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.Beat.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareBeatExcel + "\",";
                            }
                        }
                        if (NewRouteCodeslst.Count() > 0 || NewSECodelst.Count > 0)
                        {
                            fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.SalesmanRoute.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareSalesmanRouteExcel + "\",";
                        }
                        if (excelRules != null && excelRules.Contains("DayOfWeek"))
                        {
                            FormatDayOfWeek(ReadRCPFromSalesTeam);
                        }
                        if ((lstFromSalesNewOutlets != null && lstFromSalesNewOutlets.Count > 0) && excelRules != null && excelRules.Contains("CreateOutletId"))
                        {
                            CreateOutletId(ReadRCPFromDMS, lstFromSalesNewOutlets);
                            fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.Outlet.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareOutletExcel + "\",";
                        }
                        if (excelRules != null && excelRules.Contains("MobileNumber"))
                        {
                            ReadRCPFromSalesTeam.Where(record => string.IsNullOrWhiteSpace(record.OutletPhoneNo.ToString()))
                            .Select(record => { record.OutletPhoneNo = 1234567890; return record; }).ToList();
                        }
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.CPCategory.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareCPCategoryExcel + "\",";
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.CustomerRoute.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareCustomerRouteExcel + "\",";
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.BeatPlanning.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareBeatPlanningExcel + "\",";
                        if (ReadRCPFromSalesTeam.Where(x => !string.IsNullOrWhiteSpace(x.OutletIdForRemoval)).Count() > 0 ||
                            ReadRCPFromSalesTeam.Where(x => !string.IsNullOrWhiteSpace(x.OutletIdforRouteTransfer)).Count() > 0)
                        {
                            fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.OutletRemoval.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareORMExcel + "\",";
                        }
                        fileRequiredJSON = fileRequiredJSON.Remove(fileRequiredJSON.Length - 1) + "}";
                        KillExcel();
                    }
                    else
                    {
                        errorMessage = "{\"Detail\":{\"" + Common.Constants.JSON.Tags.Message.Details.Key + "\":\"" + "No valid data found." + "\"}}";
                    }
                }
            }
            catch (Exception ex)
            { }
            return fileRequiredJSON;
        }

        /// <summary>
        /// Method to create outletId
        /// </summary>
        /// <param name="ReadRCPFromDMS"></param>
        /// <param name="ReadRCPFromSalesTeam"></param>
        private void CreateOutletId(List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> lstFromSalesNewOutlets)
        {
            string tempOutletId = string.Empty;
            int counter = 1;
            try
            {
                foreach (RCPFromSalesTeamModel allFromSalesTeam in lstFromSalesNewOutlets)
                {
                    if (allFromSalesTeam.OutletCode.ToLower().Contains("new"))
                    {
                        string newOutletId = "B0" + DateTime.Now.AddYears(1).ToString("ddMMyy") + counter.ToString().PadLeft(4, '0');
                        while (CheckNewOutletIdExistance(newOutletId, lstFromSalesNewOutlets) || newOutletId == tempOutletId)
                        {
                            counter++;
                            newOutletId = "B0" + DateTime.Now.AddYears(1).ToString("ddMMyy") + counter.ToString().PadLeft(4, '0');
                        }
                        allFromSalesTeam.OutletCode = newOutletId;
                        tempOutletId = newOutletId;
                        counter++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void CreateRoutCodeAndName(string branchCode, List<RCPFromDMSModel> ReadRCPFromDMS, ref List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string tempRoutCode = string.Empty, newRoutCode = string.Empty, newRoutName = string.Empty;
            int counter = 1;
            string bCode = Enum.Parse(typeof(BranchCode), branchCode).ToString();
            try
            {
                foreach (RCPFromSalesTeamModel allFromSalesTeam in ReadRCPFromSalesTeam)
                {
                    newRoutName = bCode + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                    newRoutCode = branchCode + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                    while (CheckNewRouteCodeExistance(newRoutCode, ReadRCPFromDMS, ReadRCPFromSalesTeam) || newRoutCode == tempRoutCode)
                    {
                        counter++;
                        newRoutName = bCode + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                        newRoutCode = branchCode + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                    }
                    allFromSalesTeam.RouteCode = newRoutCode;
                    allFromSalesTeam.RouteName = newRoutName;
                    tempRoutCode = newRoutCode;
                    counter++;
                    break;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void CreateSECode(string branchCode, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, int counter)
        {
            string tempSECode = string.Empty, newSECode = string.Empty;
            string bCode = Enum.Parse(typeof(BranchCode), branchCode).ToString();
            try
            {
                newSECode = bCode + DateTime.Now.ToString("ddMMyy") + counter.ToString().PadLeft(3, '0');
                foreach (RCPFromSalesTeamModel allFromSalesTeam in ReadRCPFromSalesTeam)
                {
                    allFromSalesTeam.SECode = newSECode;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to format DayOfWeek
        /// </summary>
        /// <param name="ReadRCPFromSalesTeam"></param>
        private void FormatDayOfWeek(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                foreach (RCPFromSalesTeamModel fromSales in ReadRCPFromSalesTeam)
                {
                    string[] arrDays = Regex.Split(fromSales.DayOfWeek, "-");
                    fromSales.DayOfWeek = string.Empty;
                    foreach (string day in arrDays)
                    {
                        if (day.ToLower().Contains("m"))
                            fromSales.DayOfWeek += "M";
                        if (day.ToLower().Contains("tu"))
                            fromSales.DayOfWeek += "-Tu";
                        if (day.ToLower().Contains("w"))
                            fromSales.DayOfWeek += "-W";
                        if (day.ToLower().Contains("th"))
                            fromSales.DayOfWeek += "-Th";
                        if (day.ToLower().Contains("f"))
                            fromSales.DayOfWeek += "-F";
                        if (day.ToLower().Contains("s"))
                            fromSales.DayOfWeek += "-Sa";
                    }
                    fromSales.DayOfWeek = fromSales.DayOfWeek.StartsWith("-") ? fromSales.DayOfWeek.Substring(1) : fromSales.DayOfWeek;
                    fromSales.DayOfWeek = fromSales.DayOfWeek.EndsWith("-") ? fromSales.DayOfWeek.Substring(0, fromSales.DayOfWeek.Length - 1) : fromSales.DayOfWeek;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to populate excel file after applying the rules
        /// </summary>
        /// <param name="ReadRCPFromSalesTeam"></param>
        public void CreateFormattedExcel(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                CreateRawFileHeader(ws);
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    CreateRawFileRow(ws, ReadRCPFromSalesTeam, i);
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.FormattedRCPFromSales));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to create rows for sales team file
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="ReadRCPFromSalesTeam"></param>
        /// <param name="i"></param>
        private void CreateRawFileRow(ExcelWorksheet ws, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, int i)
        {
            ws.Cells[i + 2, 1].Value = ReadRCPFromSalesTeam[i].RegionName;
            ws.Cells[i + 2, 2].Value = ReadRCPFromSalesTeam[i].State;
            ws.Cells[i + 2, 3].Value = ReadRCPFromSalesTeam[i].City;
            ws.Cells[i + 2, 4].Value = ReadRCPFromSalesTeam[i].TownName;
            ws.Cells[i + 2, 5].Value = ReadRCPFromSalesTeam[i].ReportingTo;
            ws.Cells[i + 2, 6].Value = ReadRCPFromSalesTeam[i].SalesForceCode;
            ws.Cells[i + 2, 7].Value = ReadRCPFromSalesTeam[i].DistributedBranchCode;
            ws.Cells[i + 2, 8].Value = ReadRCPFromSalesTeam[i].RouteName;
            ws.Cells[i + 2, 9].Value = ReadRCPFromSalesTeam[i].IsNewRoute;
            ws.Cells[i + 2, 10].Value = ReadRCPFromSalesTeam[i].RouteCode;
            ws.Cells[i + 2, 11].Value = ReadRCPFromSalesTeam[i].SEType;
            ws.Cells[i + 2, 12].Value = ReadRCPFromSalesTeam[i].SECode;
            ws.Cells[i + 2, 13].Value = ReadRCPFromSalesTeam[i].IsNewSalesman;
            ws.Cells[i + 2, 14].Value = ReadRCPFromSalesTeam[i].SalesmanCategory;
            ws.Cells[i + 2, 15].Value = ReadRCPFromSalesTeam[i].DayOfWeek;
            ws.Cells[i + 2, 16].Value = ReadRCPFromSalesTeam[i].OutletCode;
            ws.Cells[i + 2, 17].Value = ReadRCPFromSalesTeam[i].OutletName;
            ws.Cells[i + 2, 18].Value = ReadRCPFromSalesTeam[i].OutletAddress;
            ws.Cells[i + 2, 19].Value = ReadRCPFromSalesTeam[i].ProductHierarchyCategoryName;
            ws.Cells[i + 2, 20].Value = ReadRCPFromSalesTeam[i].PostalCode;
            ws.Cells[i + 2, 21].Value = ReadRCPFromSalesTeam[i].Retlrtype;
            ws.Cells[i + 2, 22].Value = ReadRCPFromSalesTeam[i].CustChannelType;
            ws.Cells[i + 2, 23].Value = ReadRCPFromSalesTeam[i].CustChannelSubType;
            ws.Cells[i + 2, 24].Value = ReadRCPFromSalesTeam[i].OutletPhoneNo;
            ws.Cells[i + 2, 25].Value = ReadRCPFromSalesTeam[i].StoreType;
            ws.Cells[i + 2, 26].Value = ReadRCPFromSalesTeam[i].OutletIdForRemoval;
            ws.Cells[i + 2, 27].Value = ReadRCPFromSalesTeam[i].RouteCodeforOutletIdRemoval;
            ws.Cells[i + 2, 28].Value = ReadRCPFromSalesTeam[i].OutletIdforRouteTransfer;
            ws.Cells[i + 2, 29].Value = ReadRCPFromSalesTeam[i].RouteCodeforOutletRouteTransfer;
        }

        private void CreateSalesmanRawFileRow(ExcelWorksheet ws, List<SalesmanModel> SalesmanModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = SalesmanModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = SalesmanModelList[i].SalesmanName;
            ws.Cells[i + 2, 3].Value = SalesmanModelList[i].SalesmanName;
            ws.Cells[i + 2, 4].Value = SalesmanModelList[i].SalesmanName;
            ws.Cells[i + 2, 5].Value = SalesmanModelList[i].SEType;
            ws.Cells[i + 2, 6].Value = SalesmanModelList[i].ReportingTo;
            ws.Cells[i + 2, 7].Value = SalesmanModelList[i].ReportingLevel;
            ws.Cells[i + 2, 8].Value = SalesmanModelList[i].SalesForceCode;
            ws.Cells[i + 2, 9].Value = SalesmanModelList[i].IsActive;
            ws.Cells[i + 2, 10].Value = SalesmanModelList[i].Category;
            ws.Cells[i + 2, 11].Value = SalesmanModelList[i].JoiningDate;
            ws.Cells[i + 2, 12].Value = SalesmanModelList[i].Email;
            ws.Cells[i + 2, 13].Value = SalesmanModelList[i].PhoneNo;
        }

        private void CreateOutletRawFileRow(ExcelWorksheet ws, List<OutletModel> outletModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = outletModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = outletModelList[i].RetlrType;
            ws.Cells[i + 2, 3].Value = outletModelList[i].OutletCode;
            ws.Cells[i + 2, 4].Value = outletModelList[i].OutletName;
            ws.Cells[i + 2, 5].Value = outletModelList[i].OutletAddress1;
            ws.Cells[i + 2, 6].Value = outletModelList[i].OutletAddress2;
            ws.Cells[i + 2, 7].Value = outletModelList[i].OutletAddress3;
            ws.Cells[i + 2, 8].Value = outletModelList[i].Country;
            ws.Cells[i + 2, 9].Value = outletModelList[i].State;
            ws.Cells[i + 2, 10].Value = outletModelList[i].City;
            ws.Cells[i + 2, 11].Value = outletModelList[i].PostalCode;
            ws.Cells[i + 2, 12].Value = outletModelList[i].Email;
            ws.Cells[i + 2, 13].Value = outletModelList[i].PhoneNo;
            ws.Cells[i + 2, 14].Value = outletModelList[i].EnrollDate;
            ws.Cells[i + 2, 15].Value = outletModelList[i].TaxType;
            ws.Cells[i + 2, 16].Value = outletModelList[i].CreditBills;
            ws.Cells[i + 2, 17].Value = outletModelList[i].CreditBillAct;
            ws.Cells[i + 2, 18].Value = outletModelList[i].CustChannelType;
            ws.Cells[i + 2, 19].Value = outletModelList[i].CustChannelSubType;
            ws.Cells[i + 2, 20].Value = outletModelList[i].IsActive;
            ws.Cells[i + 2, 21].Value = outletModelList[i].CreditDaysAct;
            ws.Cells[i + 2, 22].Value = outletModelList[i].CreditLimitAct;
            ws.Cells[i + 2, 23].Value = outletModelList[i].CashDiscPerc;
            ws.Cells[i + 2, 24].Value = outletModelList[i].StoreType;
        }

        private void CreateCustomerProductCategoryRawFileRow(ExcelWorksheet ws, List<CustomerProductCategoryModel> customerProductCategoryModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = customerProductCategoryModelList[i].OutletCode;
            ws.Cells[i + 2, 2].Value = customerProductCategoryModelList[i].ProductHierarchyLevelCode;
            ws.Cells[i + 2, 3].Value = customerProductCategoryModelList[i].ProductHierarchyValueCode;
            ws.Cells[i + 2, 4].Value = customerProductCategoryModelList[i].ProductHierarchyCategoryName;
        }

        private void CreateBeatRawFileRow(ExcelWorksheet ws, List<BeatModel> beatModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = beatModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = beatModelList[i].RouteCode;
            ws.Cells[i + 2, 3].Value = beatModelList[i].RouteName;
            ws.Cells[i + 2, 4].Value = beatModelList[i].Distance;
            ws.Cells[i + 2, 5].Value = beatModelList[i].Population;
            ws.Cells[i + 2, 4].Style.Numberformat.Format = "0";
            ws.Cells[i + 2, 5].Style.Numberformat.Format = "0";
        }

        private void CreateSalsemanRouteRawFileRow(ExcelWorksheet ws, List<SalesmanRouteModel> salsemanRouteModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = salsemanRouteModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = salsemanRouteModelList[i].SalesmanCode;
            ws.Cells[i + 2, 3].Value = salsemanRouteModelList[i].RouteCode;
        }

        private void CreateCustomerRouteRawFileRow(ExcelWorksheet ws, List<CustomerRouteModel> customerRouteModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = customerRouteModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = customerRouteModelList[i].OutletCode;
            ws.Cells[i + 2, 3].Value = customerRouteModelList[i].RouteCode;
            ws.Cells[i + 2, 4].Value = customerRouteModelList[i].CoverageSequence;
        }

        private void CreateCustomerRouteRawFileRow1(ExcelWorksheet ws1, List<CustomerRouteModel> customerRouteModelList1, int i)
        {
            ws1.Cells[i + 2, 1].Value = customerRouteModelList1[i].DistributorCode;
            ws1.Cells[i + 2, 2].Value = customerRouteModelList1[i].MaximumCoverageSequenceNo;
        }

        /// <summary>
        /// Method to create header for sales team file
        /// </summary>
        /// <param name="ws"></param>
        private void CreateRawFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "Region Name";
            ws.Cells[1, 2].Value = "State";
            ws.Cells[1, 3].Value = "City";
            ws.Cells[1, 4].Value = "Town Name";
            ws.Cells[1, 5].Value = "ReportingTo";
            ws.Cells[1, 6].Value = "SalesForceCode";
            ws.Cells[1, 7].Value = "DistributedBranchCode";
            ws.Cells[1, 8].Value = "Route Name";
            ws.Cells[1, 9].Value = "IsNewRoute";
            ws.Cells[1, 10].Value = "Route Code";
            ws.Cells[1, 11].Value = "SEType";
            ws.Cells[1, 12].Value = "SECode";
            ws.Cells[1, 13].Value = "IsNewSalesman";
            ws.Cells[1, 14].Value = "SalesmanCategory";
            ws.Cells[1, 15].Value = "DayOfWeek";
            ws.Cells[1, 16].Value = "Outlet Code";
            ws.Cells[1, 17].Value = "Outlet Name";
            ws.Cells[1, 18].Value = "Outlet Address";
            ws.Cells[1, 19].Value = "ProductHierarchyCategoryName";
            ws.Cells[1, 20].Value = "PostalCode";
            ws.Cells[1, 21].Value = "Retlrtype";
            ws.Cells[1, 22].Value = "CustChannelType";
            ws.Cells[1, 23].Value = "CustChannelSubType";
            ws.Cells[1, 24].Value = "OutletPhoneNo";
            ws.Cells[1, 25].Value = "StoreType";
            ws.Cells[1, 26].Value = "OutletIdForRemoval";
            ws.Cells[1, 27].Value = "RouteCodeforOutletIdRemoval";
            ws.Cells[1, 28].Value = "OutletIdforRouteTransfer";
            ws.Cells[1, 29].Value = "RouteCodeforOutletRouteTransfer";

        }

        private void CreateSalesmanFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "SECode";
            ws.Cells[1, 3].Value = "SalesmanCode";
            ws.Cells[1, 4].Value = "SalesmanName";
            ws.Cells[1, 5].Value = "SEType";
            ws.Cells[1, 6].Value = "ReportingTo";
            ws.Cells[1, 7].Value = "ReportingLevel";
            ws.Cells[1, 8].Value = "SalesForceCode";
            ws.Cells[1, 9].Value = "IsActive";
            ws.Cells[1, 10].Value = "Category";
            ws.Cells[1, 11].Value = "JoiningDate";
            ws.Cells[1, 12].Value = "Email";
            ws.Cells[1, 13].Value = "PhoneNo";
        }

        private void CreateOutletFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "RetlrType";
            ws.Cells[1, 3].Value = "OutletCode";
            ws.Cells[1, 4].Value = "OutletName";
            ws.Cells[1, 5].Value = "OutletAddress1";
            ws.Cells[1, 6].Value = "OutletAddress2";
            ws.Cells[1, 7].Value = "OutletAddress3";
            ws.Cells[1, 8].Value = "Country";
            ws.Cells[1, 9].Value = "State";
            ws.Cells[1, 10].Value = "City";
            ws.Cells[1, 11].Value = "PostalCode";
            ws.Cells[1, 12].Value = "Email";
            ws.Cells[1, 13].Value = "PhoneNo";
            ws.Cells[1, 14].Value = "EnrollDate";
            ws.Cells[1, 15].Value = "TaxType";
            ws.Cells[1, 16].Value = "CreditBills";
            ws.Cells[1, 17].Value = "CreditBillAct";
            ws.Cells[1, 18].Value = "CustChannelType";
            ws.Cells[1, 19].Value = "CustChannelSubType";
            ws.Cells[1, 20].Value = "IsActive";
            ws.Cells[1, 21].Value = "CreditDaysAct";
            ws.Cells[1, 22].Value = "CreditLimitAct";
            ws.Cells[1, 23].Value = "CashDiscPerc";
            ws.Cells[1, 24].Value = "StoreType";
        }

        private void CreateCustomerProductCategoryFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "OutletCode";
            ws.Cells[1, 2].Value = "ProductHierarchyLevelCode";
            ws.Cells[1, 3].Value = "ProductHierarchyValueCode";
            ws.Cells[1, 4].Value = "ProductHierarchyCategoryName";
        }

        private void CreateBeatFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "RouteCode";
            ws.Cells[1, 3].Value = "RouteName";
            ws.Cells[1, 4].Value = "Distance";
            ws.Cells[1, 5].Value = "Population";
        }

        private void CreateSalsemanRouteFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "SalesmanCode";
            ws.Cells[1, 3].Value = "RouteCode";
        }

        private void CreateOutletRemovalFileHeader(ExcelWorksheet ws)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "OutletCode";
            ws.Cells[1, 3].Value = "RouteCode";
        }

        private void CreateCustomerRouteFileHeader(ExcelWorksheet ws, ExcelWorksheet ws1)
        {
            ws.Cells[1, 1].Value = "DistributorBranchCode";
            ws.Cells[1, 2].Value = "OutletCode";
            ws.Cells[1, 3].Value = "RouteCode";
            ws.Cells[1, 4].Value = "CoverageSequence";

            ws1.Cells[1, 1].Value = "DistributorCode";
            ws1.Cells[1, 2].Value = "MaximumCoverageSequenceNo";
        }

        /// <summary>
        /// Method to check if Outlet Id exists or not
        /// </summary>
        /// <param name="newOutletId"></param>
        /// <param name="ReadRCPFromDMS"></param>
        /// <param name="ReadRCPFromSalesTeam"></param>
        /// <returns></returns>
        private bool CheckNewOutletIdExistance(string newOutletId, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            var exists = (from fromSalesTeam in ReadRCPFromSalesTeam
                          where fromSalesTeam.OutletCode == newOutletId
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
        }

        private bool CheckNewSECodeExistance(string newSEId, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            var exists = (from fromSalesTeam in ReadRCPFromSalesTeam
                          where fromSalesTeam.SECode == newSEId //|| fromDMS.RouteCode == newRoutCode
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
        }

        private bool CheckNewRouteCodeExistance(string newRoutCode, List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            var exists = (from fromSalesTeam in ReadRCPFromSalesTeam
                          where fromSalesTeam.RouteCode == newRoutCode //|| fromDMS.RouteCode == newRoutCode
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
        }

        public string CreateSalesman(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string shortFileName = Common.Constants.ExcelFileName.Salesman.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.Salesman;
            KillExcel();
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            List<SalesmanModel> salesmanModelList = new List<SalesmanModel>();
            FillSalesman(salesmanModelList, ReadRCPFromSalesTeam);
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 4].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 5].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 6].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 7].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 9].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 10].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 12].EntireColumn.NumberFormat = "@";

                xlWorkSheet.Cells[2, 8].EntireColumn.NumberFormat = "0";
                xlWorkSheet.Cells[2, 13].EntireColumn.NumberFormat = "0";

                //xlWorkSheet.Cells[2, 8].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                //xlWorkSheet.Cells[2, 13].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 8].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                xlWorkSheet.Cells[2, 13].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                for (int i = 0; i < salesmanModelList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + salesmanModelList[i].DistributorBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = salesmanModelList[i].SECode;
                    xlWorkSheet.Cells[i + 2, 3].Value = salesmanModelList[i].SECode;
                    xlWorkSheet.Cells[i + 2, 4].Value = salesmanModelList[i].SECode;
                    xlWorkSheet.Cells[i + 2, 5].Value = salesmanModelList[i].SEType;
                    xlWorkSheet.Cells[i + 2, 6].Value = salesmanModelList[i].ReportingTo;
                    xlWorkSheet.Cells[i + 2, 7].Value = salesmanModelList[i].ReportingLevel;
                    xlWorkSheet.Cells[i + 2, 8].Value = salesmanModelList[i].SalesForceCode;
                    xlWorkSheet.Cells[i + 2, 9].Value = salesmanModelList[i].IsActive;
                    xlWorkSheet.Cells[i + 2, 10].Value = salesmanModelList[i].Category;
                    xlWorkSheet.Cells[i + 2, 11].Value = salesmanModelList[i].JoiningDate;
                    xlWorkSheet.Cells[i + 2, 12].Value = salesmanModelList[i].Email;
                    xlWorkSheet.Cells[i + 2, 13].Value = salesmanModelList[i].PhoneNo;
                }

                //xlWorkBook.Save();
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }

            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                KillExcel();
            }

            return shortFileName;
        }

        private void KillExcel()
        {
            var process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }

        public string CreateCustomerProductCategory(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string shortFileName = Common.Constants.ExcelFileName.CPCategory.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.CPCategory;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            List<CustomerProductCategoryModel> customerProductCategoryModelList = new List<CustomerProductCategoryModel>();
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                FillCustomerProductCategory(customerProductCategoryModelList, ReadRCPFromSalesTeam);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 4].EntireColumn.NumberFormat = "@";

                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "0";

                //xlWorkSheet.Cells[2, 2].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 2].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                for (int i = 0; i < customerProductCategoryModelList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = customerProductCategoryModelList[i].OutletCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = customerProductCategoryModelList[i].ProductHierarchyLevelCode;
                    xlWorkSheet.Cells[i + 2, 3].Value = customerProductCategoryModelList[i].ProductHierarchyValueCode;
                    xlWorkSheet.Cells[i + 2, 4].Value = customerProductCategoryModelList[i].ProductHierarchyCategoryName;
                }

                //xlWorkBook.Save();
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }

            return shortFileName;
        }

        public string CreateOutletRemoval(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string shortFileName = Common.Constants.ExcelFileName.OutletRemoval.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.OutletRemoval;
            List<OutletRouteRemovalModel> outletRouteRemovalModellList = new List<OutletRouteRemovalModel>();
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            string dBranchCode = string.Empty;
            xlApp = new Excel.Application();
            try
            {
                dBranchCode = ReadRCPFromSalesTeam.Where(x => !string.IsNullOrWhiteSpace(x.DistributedBranchCode)).Select(x => x.DistributedBranchCode).FirstOrDefault();
                outletRouteRemovalModellList = ReadRCPFromSalesTeam.Select(record => new OutletRouteRemovalModel
                {
                    DistributorBranchCode = dBranchCode,
                    OutletCode = record.OutletIdForRemoval,
                    RouteCode = record.RouteCodeforOutletIdRemoval
                }).ToList();

                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";
                for (int i = 0; i < outletRouteRemovalModellList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + outletRouteRemovalModellList[i].DistributorBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = outletRouteRemovalModellList[i].OutletCode;
                    xlWorkSheet.Cells[i + 2, 3].Value = "'" + outletRouteRemovalModellList[i].RouteCode;
                }

                //xlWorkBook.Save();
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }

            return shortFileName;
        }

        public string CreateBeat(List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string shortFileName = Common.Constants.ExcelFileName.Beat.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.Beat;
            KillExcel();
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                var rcpBeat = ReadRCPFromSalesTeam.Where(x => x.IsNewRoute.ToLower().Equals("new")).GroupBy(x => new
                {
                    x.DistributedBranchCode,
                    x.RouteName,
                    x.RouteCode


                }).Select(grp => grp.Key).ToList();

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 4].EntireColumn.NumberFormat = "0";
                xlWorkSheet.Cells[2, 5].EntireColumn.NumberFormat = "0";
                //xlWorkSheet.Cells[2, 4].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                //xlWorkSheet.Cells[2, 5].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 4].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                xlWorkSheet.Cells[2, 5].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                for (int i = 0; i < rcpBeat.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + rcpBeat[i].DistributedBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = "'" + rcpBeat[i].RouteCode;
                    xlWorkSheet.Cells[i + 2, 3].Value = rcpBeat[i].RouteName;
                    xlWorkSheet.Cells[i + 2, 4].Value = 50;
                    xlWorkSheet.Cells[i + 2, 5].Value = 50;
                }

                //xlWorkBook.Save();
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            }
            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }
            return shortFileName;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public string CreateSalesmanRoute(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string shortFileName = Common.Constants.ExcelFileName.SalesmanRoute.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.SalesmanRoute;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            List<SalesmanRouteModel> salesmanRouteModelList = new List<SalesmanRouteModel>();
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                FillSalsemanRoute(salesmanRouteModelList, ReadRCPFromSalesTeam);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";

                for (int i = 0; i < salesmanRouteModelList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + salesmanRouteModelList[i].DistributorBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = salesmanRouteModelList[i].SalesmanCode;
                    xlWorkSheet.Cells[i + 2, 3].Value = "'" + salesmanRouteModelList[i].RouteCode;
                }
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                // xlWorkBook.Save();
            }
            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }

            return shortFileName;
        }

        public string CreateCustomerRoute(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, long processInstanceId, List<RCPFromSalesTeamModel> lstOutletRouteTransfer)
        {
            List<CustomerRouteModel> customerRouteModelList = new List<CustomerRouteModel>();
            List<CustomerRouteModel> tempcustomerRouteModelList = new List<CustomerRouteModel>();
            CustomerRouteModel tempRcpFromSalesTeamModel = null;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            string fileName = string.Empty;
            string sheetName = "";
            try
            {
                customerRouteModelList = ReadCustomerRoutefromDMS(processInstanceId);
                long MaxCoverSeq = 0;
                List<string> DistCode = ReadRCPFromSalesTeam.Select(x => x.DistributedBranchCode).Distinct().ToList();
                foreach (string item in DistCode)
                {
                    MaxCoverSeq = Convert.ToInt64(customerRouteModelList.Where(x => x.DistributorCode.Equals(item)).
                        OrderByDescending(x => x.MaximumCoverageSequenceNo).Select(x => x.MaximumCoverageSequenceNo).FirstOrDefault());
                    foreach (var fromSales in ReadRCPFromSalesTeam)
                    {
                        tempRcpFromSalesTeamModel = new CustomerRouteModel();
                        if (fromSales.DistributedBranchCode.Equals(item))
                        {
                            tempRcpFromSalesTeamModel.DistributorBranchCode = fromSales.DistributedBranchCode;
                            tempRcpFromSalesTeamModel.OutletCode = fromSales.OutletCode;
                            tempRcpFromSalesTeamModel.RouteCode = fromSales.RouteCode;
                            tempRcpFromSalesTeamModel.CoverageSequence = (++MaxCoverSeq).ToString();
                            tempcustomerRouteModelList.Add(tempRcpFromSalesTeamModel);
                        }
                        else { }
                    }
                }
                GetFileName(processInstanceId, ref fileName, ref sheetName, Common.Constants.Process.States.RCPExccelProcess.CustomerRouteDownloadCompleted);


                xlWorkBook = xlApp.Workbooks.Open(this.folderPath + fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";

                xlWorkSheet.Cells[2, 4].EntireColumn.NumberFormat = "0";
                //xlWorkSheet.Cells[2, 4].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 4].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                for (int i = 0; i < tempcustomerRouteModelList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + tempcustomerRouteModelList[i].DistributorBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = tempcustomerRouteModelList[i].OutletCode;
                    xlWorkSheet.Cells[i + 2, 3].Value = "'" + tempcustomerRouteModelList[i].RouteCode;
                    xlWorkSheet.Cells[i + 2, 4].Value = tempcustomerRouteModelList[i].CoverageSequence;
                }
                MaxCoverSeq = Convert.ToInt64(tempcustomerRouteModelList[tempcustomerRouteModelList.Count - 1].CoverageSequence);
                for (int i = 0; i < lstOutletRouteTransfer.Count; i++)
                {
                    MaxCoverSeq++;
                    xlWorkSheet.Cells[tempcustomerRouteModelList.Count + i + 2, 1].Value = "'" + tempcustomerRouteModelList[0].DistributorBranchCode;
                    xlWorkSheet.Cells[tempcustomerRouteModelList.Count + i + 2, 2].Value = lstOutletRouteTransfer[i].OutletIdforRouteTransfer;
                    xlWorkSheet.Cells[tempcustomerRouteModelList.Count + i + 2, 3].Value = "'" + lstOutletRouteTransfer[i].RouteCodeforOutletRouteTransfer;
                    xlWorkSheet.Cells[tempcustomerRouteModelList.Count + i + 2, 4].Value = MaxCoverSeq;
                    
                }
                xlWorkBook.Save();
                //xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }
            return fileName;
        }

        public List<BeatPlanningModel> FillBeatPlanningList(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            List<BeatPlanningModel> beatPlanningModelList = new List<BeatPlanningModel>();
            BeatPlanningModel beatPlanningModel = null;
            try
            {
                var tempRCPFromSalesTeam = (from c in ReadRCPFromSalesTeam
                                            group c by new { c.DistributedBranchCode, c.DayOfWeek, c.RouteName, c.SECode, c.ReportingTo, c.RouteCode } into grp
                                            select new
                                            {
                                                grp.Key.DistributedBranchCode,
                                                grp.Key.DayOfWeek,
                                                grp.Key.RouteName,
                                                grp.Key.SECode,
                                                grp.Key.ReportingTo,
                                                grp.Key.RouteCode,


                                            }).Distinct().ToList();
                for (int i = 0; i < tempRCPFromSalesTeam.Count; i++)
                {
                    beatPlanningModel = new BeatPlanningModel();
                    beatPlanningModel.SECode = tempRCPFromSalesTeam[i].SECode;
                    beatPlanningModel.WDCode = tempRCPFromSalesTeam[i].DistributedBranchCode;
                    beatPlanningModel.RouteName = tempRCPFromSalesTeam[i].RouteName;
                    beatPlanningModel.VisitDates = GetVisitDatesFromDaysOfweek(tempRCPFromSalesTeam[i].DayOfWeek);
                    beatPlanningModel.EndDate = GetEndDateFromDaysOfweek(tempRCPFromSalesTeam[i].DayOfWeek);
                    beatPlanningModelList.Add(beatPlanningModel);
                }
            }
            catch (Exception)
            {

            }
            return beatPlanningModelList;
        }

        public List<BeatPlanningModel> CreateBeatPlanning(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, ref string fileName)
        {
            string json = string.Empty;
            List<BeatPlanningModel> beatPlanningModelList = new List<BeatPlanningModel>();
            try
            {
                var tempRCPFromSalesTeam = (from c in ReadRCPFromSalesTeam
                                            group c by new { c.DistributedBranchCode, c.DayOfWeek, c.RouteName, c.SECode, c.ReportingTo, c.RouteCode } into grp
                                            select new
                                            {
                                                grp.Key.DistributedBranchCode,
                                                grp.Key.DayOfWeek,
                                                grp.Key.RouteName,
                                                grp.Key.SECode,
                                                grp.Key.ReportingTo,
                                                grp.Key.RouteCode,


                                            }).Distinct().ToList();

                beatPlanningModelList = FillBeatPlanningList(ReadRCPFromSalesTeam);
                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "SE Code";
                ws.Cells[1, 2].Value = "Day Of Week";
                ws.Cells[1, 3].Value = "WD Code";
                ws.Cells[1, 4].Value = "AM Code";
                ws.Cells[1, 5].Value = "Actual Route Code";
                ws.Cells[1, 6].Value = "Final Route Name";

                for (int i = 0; i < tempRCPFromSalesTeam.Count; i++)
                {
                    ws.Cells[i + 2, 1].Value = tempRCPFromSalesTeam[i].SECode;
                    ws.Cells[i + 2, 2].Value = tempRCPFromSalesTeam[i].DayOfWeek;
                    ws.Cells[i + 2, 3].Value = tempRCPFromSalesTeam[i].DistributedBranchCode;
                    ws.Cells[i + 2, 4].Value = tempRCPFromSalesTeam[i].ReportingTo;
                    ws.Cells[i + 2, 5].Value = tempRCPFromSalesTeam[i].RouteCode;
                    ws.Cells[i + 2, 6].Value = tempRCPFromSalesTeam[i].RouteName;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                fileName = Common.Constants.ExcelFileName.BeatPlanning.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + fileName));
            }
            catch (Exception ex)
            {
                throw;
            }
            string fName = "BeatPlanning.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
            return beatPlanningModelList;
        }

        private string GetVisitDatesFromDaysOfweek(string dayOfWeek)
        {
            string result = string.Empty;
            try
            {
                int currentDayOfWeek = Convert.ToInt32(DateTime.Now.DayOfWeek);
                string[] arrDays = Regex.Split(dayOfWeek, "-");
                foreach (string day in arrDays)
                {
                    int dayId = GetDayIdFromDay(day);
                    if (dayId < currentDayOfWeek)
                        result += DateTime.Now.AddDays(7 - currentDayOfWeek + dayId).ToString("dd/MM/yyyy").Replace('-', '/') + ",";
                    else
                        result += DateTime.Now.AddDays(dayId - currentDayOfWeek).ToString("dd/MM/yyyy").Replace('-', '/') + ",";
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return result.EndsWith(",") ? result.Substring(0, result.Length - 1) : result;
        }

        private string GetEndDateFromDaysOfweek(string dayOfWeek)
        {
            string charDay, result, dow = string.Empty;
            try
            {
                DateTime year = DateTime.Now.AddYears(1);
                DateTime date = new DateTime(year.Year, 12, 31);
                dow = date.DayOfWeek.ToString();
                charDay = (dow.Equals("Thursday") || dow.Equals("Tuesday") || dow.Equals("Saturday")) ? dow.Substring(0, 2) : dow.Substring(0, 1);
                int i = 1;
                while (!dayOfWeek.Contains(charDay))
                {
                    dow = date.AddDays(-1).DayOfWeek.ToString();
                    charDay = (dow.Equals("Thursday") || dow.Equals("Tuesday") || dow.Equals("Saturday")) ? dow.Substring(0, 2) : dow.Substring(0, 1);
                    date = date.AddDays(-1);
                    i++;
                }
                result = date.ToString("dd/MM/yyyy").Replace('-', '/');
            }
            catch (Exception ex)
            {
                result = string.Empty;
                throw;
            }
            return result;
        }

        private int GetDayIdFromDay(string day)
        {
            int result = 0;
            switch (day)
            {
                case "M": result = 1; break;
                case "Tu": result = 2; break;
                case "W": result = 3; break;
                case "Th": result = 4; break;
                case "F": result = 5; break;
                case "Sa": result = 6; break;
            }
            return result;
        }

        private void FillSalesman(List<SalesmanModel> SalesmanModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            SalesmanModel salesmanModel = null;
            try
            {
                var templst = ReadRCPFromSalesTeam.Where(x => x.IsNewSalesman.ToLower().Equals("new")).GroupBy(x => new { x.SECode }).Select(grp => grp.ToList()).ToList();
                List<RCPFromSalesTeamModel> lstFromSalesNewSEId = new List<RCPFromSalesTeamModel>();
                foreach (var item in templst)
                {
                    foreach (RCPFromSalesTeamModel fromSales in item)
                    {
                        lstFromSalesNewSEId.Add(fromSales);
                        break;
                    }
                }
                for (int i = 0; i < lstFromSalesNewSEId.Count; i++)
                {
                    salesmanModel = new SalesmanModel();
                    salesmanModel.DistributorBranchCode = lstFromSalesNewSEId[i].DistributedBranchCode;
                    salesmanModel.SECode = lstFromSalesNewSEId[i].SECode;
                    salesmanModel.SalesmanCode = lstFromSalesNewSEId[i].SECode;
                    salesmanModel.SalesmanName = lstFromSalesNewSEId[i].SECode;
                    salesmanModel.SEType = lstFromSalesNewSEId[i].SEType;
                    salesmanModel.ReportingTo = lstFromSalesNewSEId[i].ReportingTo;
                    salesmanModel.ReportingLevel = "Assistant Manager";
                    salesmanModel.SalesForceCode = lstFromSalesNewSEId[i].SalesForceCode;
                    salesmanModel.IsActive = "Y";
                    salesmanModel.Category = lstFromSalesNewSEId[i].SalesmanCategory;
                    salesmanModel.JoiningDate = DateTime.Now.ToString("dd/MM/yyyy").Replace("-", "/");
                    salesmanModel.Email = "abc@gmail.com";
                    salesmanModel.PhoneNo = 1234567890;
                    SalesmanModelList.Add(salesmanModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillOutlet(List<OutletModel> outletModelList, List<RCPFromSalesTeamModel> lstFromSalesNewOutlets)
        {
            OutletModel outletModel = null;
            try
            {
                for (int i = 0; i < lstFromSalesNewOutlets.Count; i++)
                {
                    outletModel = new OutletModel();
                    outletModel.CashDiscPerc = 0;
                    outletModel.City = lstFromSalesNewOutlets[i].City;
                    outletModel.Country = "India";
                    outletModel.CreditBillAct = "None";
                    outletModel.CreditBills = 0;
                    outletModel.CreditDaysAct = "None";
                    outletModel.CreditLimitAct = "None";
                    outletModel.CustChannelSubType = lstFromSalesNewOutlets[i].CustChannelSubType;
                    outletModel.CustChannelType = lstFromSalesNewOutlets[i].CustChannelType;
                    outletModel.DistributorBranchCode = lstFromSalesNewOutlets[i].DistributedBranchCode;
                    outletModel.Email = "abc@gmail.com";
                    outletModel.EnrollDate = string.Empty;
                    outletModel.IsActive = "Y";
                    outletModel.OutletAddress1 = Regex.Replace(lstFromSalesNewOutlets[i].OutletAddress, @"[^0-9a-zA-Z]+", " ");
                    outletModel.OutletAddress2 = string.Empty;
                    outletModel.OutletAddress3 = string.Empty;
                    outletModel.OutletCode = lstFromSalesNewOutlets[i].OutletCode;
                    outletModel.OutletName = Regex.Replace(lstFromSalesNewOutlets[i].OutletName, @"[^0-9a-zA-Z]+", " ");
                    outletModel.PhoneNo = string.IsNullOrWhiteSpace(lstFromSalesNewOutlets[i].OutletPhoneNo.ToString()) ? 1234567890 : lstFromSalesNewOutlets[i].OutletPhoneNo;
                    outletModel.PostalCode = lstFromSalesNewOutlets[i].PostalCode.ToString().Length == 6 ? lstFromSalesNewOutlets[i].PostalCode : 123456;
                    outletModel.RetlrType = string.IsNullOrWhiteSpace(lstFromSalesNewOutlets[i].Retlrtype) ? "Retailer" : lstFromSalesNewOutlets[i].Retlrtype;
                    outletModel.State = lstFromSalesNewOutlets[i].State;
                    outletModel.StoreType = lstFromSalesNewOutlets[i].StoreType;
                    outletModel.TaxType = "VAT";
                    outletModelList.Add(outletModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillCustomerProductCategory(List<CustomerProductCategoryModel> customerProductCategoryModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            CustomerProductCategoryModel customerProductCategoryModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    customerProductCategoryModel = new CustomerProductCategoryModel();
                    customerProductCategoryModel.OutletCode = ReadRCPFromSalesTeam[i].OutletCode;
                    customerProductCategoryModel.ProductHierarchyCategoryName = ReadRCPFromSalesTeam[i].ProductHierarchyCategoryName;
                    customerProductCategoryModel.ProductHierarchyLevelCode = 100;
                    customerProductCategoryModel.ProductHierarchyValueCode = "CIG";
                    customerProductCategoryModelList.Add(customerProductCategoryModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillOutletRemoval(List<OutletRouteRemovalModel> outletRemovalModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            OutletRouteRemovalModel outletRemovalModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    outletRemovalModel = new OutletRouteRemovalModel();
                    outletRemovalModel.OutletCode = ReadRCPFromSalesTeam[i].OutletCode;
                    outletRemovalModel.RouteCode = ReadRCPFromSalesTeam[i].RouteCode;
                    outletRemovalModel.DistributorBranchCode = ReadRCPFromSalesTeam[i].DistributedBranchCode;
                    outletRemovalModelList.Add(outletRemovalModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillBeat(List<BeatModel> BeatModelList, List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            BeatModel beatModel = null;
            try
            {
                var tempBeatModelList = ReadRCPFromSalesTeam.GroupBy(x => new
                {
                    x.DistributedBranchCode,
                    x.RouteCode,
                    x.RouteName

                }).Select(grp => grp.Key).ToList();

                foreach (var item in tempBeatModelList)
                {
                    beatModel = new BeatModel();
                    beatModel.DistributorBranchCode = item.DistributedBranchCode;
                    beatModel.RouteCode = item.RouteCode;
                    beatModel.RouteName = item.RouteName;
                    beatModel.Distance = 50;
                    beatModel.Population = 50;
                    BeatModelList.Add(beatModel);
                }

            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillSalsemanRoute(List<SalesmanRouteModel> salesmanRouteModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            SalesmanRouteModel salesmanRouteModel = null;
            try
            {
                var tempsalesmanRouteModelList = ReadRCPFromSalesTeam.Where(x => x.IsNewRoute.ToLower().Equals("new") || x.IsNewSalesman.ToLower().Equals("new"))
                    .GroupBy(x => new
                    {
                        x.DistributedBranchCode,
                        x.SECode,
                        x.RouteCode

                    }).Select(grp => grp.Key).ToList();
                foreach (var item in tempsalesmanRouteModelList)
                {
                    salesmanRouteModel = new SalesmanRouteModel();
                    salesmanRouteModel.DistributorBranchCode = item.DistributedBranchCode;
                    salesmanRouteModel.SalesmanCode = item.SECode;
                    salesmanRouteModel.RouteCode = item.RouteCode;
                    salesmanRouteModelList.Add(salesmanRouteModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void FillOutletRouteRemoval(List<OutletRouteRemovalModel> outletRouteRemovalModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            OutletRouteRemovalModel outletRouteRemovalModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    outletRouteRemovalModel = new OutletRouteRemovalModel();
                    outletRouteRemovalModel.DistributorBranchCode = ReadRCPFromSalesTeam[i].DistributedBranchCode;
                    outletRouteRemovalModel.OutletCode = ReadRCPFromSalesTeam[i].OutletCode;
                    outletRouteRemovalModel.RouteCode = ReadRCPFromSalesTeam[i].RouteCode;
                    outletRouteRemovalModelList.Add(outletRouteRemovalModel);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public List<CustomerRouteModel> ReadCustomerRoutefromDMS(long processInstanceId)
        {
            List<CustomerRouteModel> cutomerRoutefromDMS = new List<CustomerRouteModel>();
            DataTable data = null;
            string fileName = string.Empty;
            string sheetName = "";
            try
            {

                GetFileName(processInstanceId, ref fileName, ref sheetName, Common.Constants.Process.States.RCPExccelProcess.CustomerRouteDownloadCompleted);
                bool isOlderExcel = IsOlderVersionExcelNew(fileName);
                if (isOlderExcel)
                {
                    fileName = this.folderPath + fileName;
                    var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", fileName);
                    var adapter = new OleDbDataAdapter("SELECT * FROM [Details sheet$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "FromDMS");
                    data = ds.Tables["FromDMS"];
                    cutomerRoutefromDMS = ConvertDataTable<CustomerRouteModel>(data);
                }
                else
                {
                    fileName = this.folderPath + fileName;
                    FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
                    DataTable dTable = new DataTable();
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets["Details sheet"];
                        int totalRows = workSheet.Dimension.Rows;

                        for (int i = 0; i <= totalRows; i++)
                        {
                            cutomerRoutefromDMS.Add(new CustomerRouteModel
                            {
                                DistributorCode = workSheet.Cells[i + 2, 1].Value != null ? workSheet.Cells[i + 2, 1].Value.ToString() : string.Empty,
                                MaximumCoverageSequenceNo = workSheet.Cells[i + 2, 2].Value != null ? workSheet.Cells[i + 2, 2].Value.ToString() : string.Empty
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return cutomerRoutefromDMS;
        }

        public string CreateSalesmanFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<SalesmanModel> rcpSalesman = ReadRCPFromSalesTeam.Select(e => new SalesmanModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    SECode = e.SECode,
                    SalesmanCode = e.SECode,
                    SalesmanName = e.SECode,
                    SEType = e.SEType,
                    ReportingTo = e.ReportingTo,
                    ReportingLevel = "Assistant Manager",
                    SalesForceCode = e.SalesForceCode,
                    IsActive = "Y",
                    Category = e.SalesmanCategory,
                    JoiningDate = " ",
                    Email = "abc@gmail.com",
                    PhoneNo = 1234567890
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "DistributorBranchCode";
                ws.Cells[1, 2].Value = "SECode";
                ws.Cells[1, 3].Value = "SalesmanCode";
                ws.Cells[1, 4].Value = "SalesmanName";
                ws.Cells[1, 5].Value = "SEType";
                ws.Cells[1, 6].Value = "ReportingTo";
                ws.Cells[1, 7].Value = "ReportingLevel";
                ws.Cells[1, 8].Value = "SalesForceCode";
                ws.Cells[1, 9].Value = "IsActive";
                ws.Cells[1, 10].Value = "Category";
                ws.Cells[1, 11].Value = "JoiningDate";
                ws.Cells[1, 12].Value = "Email";
                ws.Cells[1, 13].Value = "PhoneNo";

                for (int i = 2; i < rcpSalesman.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpSalesman[i].DistributorBranchCode;
                    ws.Cells[i, 2].Value = rcpSalesman[i].SECode;
                    ws.Cells[i, 3].Value = rcpSalesman[i].SalesmanCode;
                    ws.Cells[i, 4].Value = rcpSalesman[i].SalesmanName;
                    ws.Cells[i, 5].Value = rcpSalesman[i].SEType;
                    ws.Cells[i, 6].Value = rcpSalesman[i].ReportingTo;
                    ws.Cells[i, 7].Value = rcpSalesman[i].ReportingLevel;
                    ws.Cells[i, 8].Value = rcpSalesman[i].SalesForceCode;
                    ws.Cells[i, 9].Value = rcpSalesman[i].IsActive;
                    ws.Cells[i, 10].Value = rcpSalesman[i].Category;
                    ws.Cells[i, 11].Value = rcpSalesman[i].Email;
                    ws.Cells[i, 12].Value = rcpSalesman[i].PhoneNo;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.Salesman));
            }
            catch (Exception ex)
            {
                throw;
            }
            return this.folderPath + Common.Constants.ExcelFileName.Salesman;
        }

        public void CreateCustomerRouteFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<CustomerRouteModel> rcpCustomerRoute = ReadRCPFromSalesTeam.Select(e => new CustomerRouteModel
                {
                    DistributorCode = e.DistributedBranchCode,
                    OutletCode = e.OutletCode,
                    RouteCode = e.RouteCode,
                    MaximumCoverageSequenceNo = " "
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "DistributorBranchCode";
                ws.Cells[1, 2].Value = "OutletCode";
                ws.Cells[1, 3].Value = "RouteCode";
                ws.Cells[1, 4].Value = "CoverageSequence";

                for (int i = 2; i < rcpCustomerRoute.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpCustomerRoute[i].DistributorCode;
                    ws.Cells[i, 2].Value = rcpCustomerRoute[i].OutletCode;
                    ws.Cells[i, 3].Value = rcpCustomerRoute[i].RouteCode;
                    ws.Cells[i, 4].Value = rcpCustomerRoute[i].MaximumCoverageSequenceNo;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.CustomerRoute));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public string CreateOutlet(List<RCPFromSalesTeamModel> lstFromSalesNewOutlets)
        {
            string shortFileName = Common.Constants.ExcelFileName.Outlet.Replace(".xls", "") + DateTime.Now.ToString("ddMMyyhhmmss") + ".xls";
            string fileName = this.folderPath + Common.Constants.ExcelFileName.Outlet;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            List<OutletModel> outletModelList = new List<OutletModel>();
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                FillOutlet(outletModelList, lstFromSalesNewOutlets);

                xlWorkSheet.Cells[2, 1].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 2].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 3].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 4].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 5].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 6].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 7].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 8].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 9].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 10].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 12].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 14].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 15].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 17].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 18].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 19].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 20].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 21].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 22].EntireColumn.NumberFormat = "@";
                xlWorkSheet.Cells[2, 24].EntireColumn.NumberFormat = "@";

                xlWorkSheet.Cells[2, 11].EntireColumn.NumberFormat = "0";
                xlWorkSheet.Cells[2, 13].EntireColumn.NumberFormat = "0";
                xlWorkSheet.Cells[2, 16].EntireColumn.NumberFormat = "0";
                xlWorkSheet.Cells[2, 23].EntireColumn.NumberFormat = "0";

                xlWorkSheet.Cells[2, 14].EntireColumn.NumberFormat = "MM/DD/YYYY";


                //xlWorkSheet.Cells[2, 11].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 11].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                // xlWorkSheet.Cells[2, 13].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 13].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                //xlWorkSheet.Cells[2, 16].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 16].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                //xlWorkSheet.Cells[2, 23].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                xlWorkSheet.Cells[2, 23].EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                for (int i = 0; i < outletModelList.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1].Value = "'" + outletModelList[i].DistributorBranchCode;
                    xlWorkSheet.Cells[i + 2, 2].Value = outletModelList[i].RetlrType;
                    xlWorkSheet.Cells[i + 2, 3].Value = outletModelList[i].OutletCode;
                    xlWorkSheet.Cells[i + 2, 4].Value = outletModelList[i].OutletName;
                    xlWorkSheet.Cells[i + 2, 5].Value = outletModelList[i].OutletAddress1.Length > 50 ? outletModelList[i].OutletAddress1.Substring(0, 44) : outletModelList[i].OutletAddress1;
                    xlWorkSheet.Cells[i + 2, 6].Value = outletModelList[i].OutletAddress2;
                    xlWorkSheet.Cells[i + 2, 7].Value = outletModelList[i].OutletAddress3;
                    xlWorkSheet.Cells[i + 2, 8].Value = outletModelList[i].Country;
                    xlWorkSheet.Cells[i + 2, 9].Value = outletModelList[i].State;
                    xlWorkSheet.Cells[i + 2, 10].Value = outletModelList[i].City;
                    xlWorkSheet.Cells[i + 2, 11].Value = outletModelList[i].PostalCode;
                    xlWorkSheet.Cells[i + 2, 12].Value = outletModelList[i].Email;
                    xlWorkSheet.Cells[i + 2, 13].Value = outletModelList[i].PhoneNo;
                    xlWorkSheet.Cells[i + 2, 14].Value = outletModelList[i].EnrollDate;
                    xlWorkSheet.Cells[i + 2, 15].Value = outletModelList[i].TaxType;
                    xlWorkSheet.Cells[i + 2, 16].Value = outletModelList[i].CreditBills;
                    xlWorkSheet.Cells[i + 2, 17].Value = outletModelList[i].CreditBillAct;
                    xlWorkSheet.Cells[i + 2, 18].Value = outletModelList[i].CustChannelType;
                    xlWorkSheet.Cells[i + 2, 19].Value = outletModelList[i].CustChannelSubType;
                    xlWorkSheet.Cells[i + 2, 20].Value = outletModelList[i].IsActive;
                    xlWorkSheet.Cells[i + 2, 21].Value = outletModelList[i].CreditDaysAct;
                    xlWorkSheet.Cells[i + 2, 22].Value = outletModelList[i].CreditLimitAct;
                    xlWorkSheet.Cells[i + 2, 23].Value = outletModelList[i].CashDiscPerc;
                    xlWorkSheet.Cells[i + 2, 24].Value = outletModelList[i].StoreType;
                }

                //xlWorkBook.Save();
                xlWorkBook.SaveAs(this.folderPath + shortFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            { }
            finally
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                KillExcel();
            }
            return shortFileName;
        }

        public void CreateOutletFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<OutletModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new OutletModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    RetlrType = e.Retlrtype,
                    OutletCode = e.OutletCode,
                    OutletName = e.OutletName,
                    OutletAddress1 = e.OutletAddress,
                    Country = "India",
                    State = e.State,
                    City = e.City,
                    PostalCode = e.PostalCode,
                    Email = "abc@gmail.com",
                    PhoneNo = 1234567890,
                    EnrollDate = string.Empty,
                    TaxType = "VAT",
                    CreditBills = 0,
                    CreditBillAct = string.Empty,
                    CustChannelType = e.CustChannelType,
                    CustChannelSubType = e.CustChannelSubType,
                    IsActive = "Yes",
                    CreditDaysAct = "None",
                    CreditLimitAct = "None",
                    CashDiscPerc = 0,
                    StoreType = e.StoreType
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "DistributorBranchCode";
                ws.Cells[1, 2].Value = "RetlrType";
                ws.Cells[1, 3].Value = "OutletCode";
                ws.Cells[1, 4].Value = "OutletName";
                ws.Cells[1, 5].Value = "OutletAddress1";
                ws.Cells[1, 6].Value = "OutletAddress2";
                ws.Cells[1, 7].Value = "OutletAddress3";
                ws.Cells[1, 8].Value = "Country";
                ws.Cells[1, 9].Value = "State";
                ws.Cells[1, 10].Value = "City";
                ws.Cells[1, 11].Value = "PostalCode";
                ws.Cells[1, 12].Value = "Email";
                ws.Cells[1, 13].Value = "PhoneNo";
                ws.Cells[1, 14].Value = "EnrollDate";
                ws.Cells[1, 15].Value = "TaxType";
                ws.Cells[1, 16].Value = "CreditBills";
                ws.Cells[1, 17].Value = "CreditBillAct";
                ws.Cells[1, 18].Value = "CustChannelType";
                ws.Cells[1, 19].Value = "CustChannelSubType";
                ws.Cells[1, 20].Value = "IsActive";
                ws.Cells[1, 21].Value = "CreditDaysAct";
                ws.Cells[1, 22].Value = "CreditLimitAct";
                ws.Cells[1, 23].Value = "CashDiscPerc";
                ws.Cells[1, 24].Value = "StoreType";

                for (int i = 2; i < rcpOutlet.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpOutlet[i].DistributorBranchCode;
                    ws.Cells[i, 2].Value = rcpOutlet[i].RetlrType;
                    ws.Cells[i, 3].Value = rcpOutlet[i].OutletCode;
                    ws.Cells[i, 4].Value = rcpOutlet[i].OutletName;
                    ws.Cells[i, 5].Value = rcpOutlet[i].OutletAddress1;
                    ws.Cells[i, 6].Value = rcpOutlet[i].OutletAddress2;
                    ws.Cells[i, 7].Value = rcpOutlet[i].OutletAddress3;
                    ws.Cells[i, 8].Value = rcpOutlet[i].Country;
                    ws.Cells[i, 9].Value = rcpOutlet[i].State;
                    ws.Cells[i, 10].Value = rcpOutlet[i].City;
                    ws.Cells[i, 11].Value = rcpOutlet[i].PostalCode;
                    ws.Cells[i, 12].Value = rcpOutlet[i].Email;
                    ws.Cells[1, 13].Value = rcpOutlet[i].PhoneNo;
                    ws.Cells[1, 14].Value = rcpOutlet[i].EnrollDate;
                    ws.Cells[1, 15].Value = rcpOutlet[i].TaxType;
                    ws.Cells[1, 16].Value = rcpOutlet[i].CreditBills;
                    ws.Cells[1, 17].Value = rcpOutlet[i].CreditBillAct;
                    ws.Cells[1, 18].Value = rcpOutlet[i].CustChannelType;
                    ws.Cells[1, 19].Value = rcpOutlet[i].CustChannelSubType;
                    ws.Cells[1, 20].Value = rcpOutlet[i].IsActive;
                    ws.Cells[1, 21].Value = rcpOutlet[i].CreditDaysAct;
                    ws.Cells[1, 22].Value = rcpOutlet[i].CreditLimitAct;
                    ws.Cells[1, 23].Value = rcpOutlet[i].CashDiscPerc;
                    ws.Cells[1, 24].Value = rcpOutlet[i].StoreType;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.Outlet));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void CreateSalesmanRouteFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<SalesmanRouteModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new SalesmanRouteModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    SalesmanCode = e.SECode,
                    RouteCode = e.RouteCode
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "DistributorBranchCode";
                ws.Cells[1, 2].Value = "SalesmanCode";
                ws.Cells[1, 3].Value = "RouteCode";

                for (int i = 2; i < rcpOutlet.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpOutlet[i].DistributorBranchCode;
                    ws.Cells[i, 2].Value = rcpOutlet[i].SalesmanCode;
                    ws.Cells[i, 3].Value = rcpOutlet[i].RouteCode;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.SalesmanRoute));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void CreateCustomerProductCategoryFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<CustomerProductCategoryModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new CustomerProductCategoryModel
                {
                    OutletCode = e.OutletCode,
                    ProductHierarchyLevelCode = 100,
                    ProductHierarchyValueCode = "CIG",
                    ProductHierarchyCategoryName = e.ProductHierarchyCategoryName
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "OutletCode";
                ws.Cells[1, 2].Value = "ProductHierarchyLevelCode";
                ws.Cells[1, 3].Value = "ProductHierarchyValueCode";
                ws.Cells[1, 3].Value = "ProductHierarchyCategoryName";

                for (int i = 0; i < rcpOutlet.Count; i++)
                {
                    ws.Cells[i + 2, 1].Value = rcpOutlet[i].OutletCode;
                    ws.Cells[i + 2, 2].Value = rcpOutlet[i].ProductHierarchyLevelCode;
                    ws.Cells[i + 2, 3].Value = rcpOutlet[i].ProductHierarchyValueCode;
                    ws.Cells[i + 2, 4].Value = rcpOutlet[i].ProductHierarchyCategoryName;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.CPCategory));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void CreateBeatFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<BeatModel> rcpBeat = ReadRCPFromSalesTeam.Select(e => new BeatModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    RouteCode = e.RouteCode,
                    RouteName = e.RouteName,
                    Distance = 50,
                    Population = 50,
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "DistributorBranchCode";
                ws.Cells[1, 2].Value = "RouteCode";
                ws.Cells[1, 3].Value = "RouteName";
                ws.Cells[1, 4].Value = "Distance";
                ws.Cells[1, 5].Value = "Population";

                for (int i = 2; i < rcpBeat.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpBeat[i].DistributorBranchCode;
                    ws.Cells[i, 2].Value = rcpBeat[i].RouteCode;
                    ws.Cells[i, 3].Value = rcpBeat[i].RouteName;
                    ws.Cells[i, 4].Value = rcpBeat[i].Distance;
                    ws.Cells[i, 5].Value = rcpBeat[i].Population;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.Beat));
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void GetFileName(long processInstanceId, ref string fileName, ref string sheetName, string stateID)
        {
            string metadata = string.Empty;
            using (GPILEntities entities = new GPILEntities())
            {
                TBL_ProcessInstanceDetails tblProcessInstance = entities.TBL_ProcessInstanceDetails.Where(x => x.ProcessInstanceId == processInstanceId && x.StateId == stateID).OrderByDescending(x => x.SequenceId).FirstOrDefault();
                metadata = entities.TBL_ProcessInstanceData.Where(x => x.ProcessInstanceId == tblProcessInstance.ProcessInstanceId && x.SequenceId == tblProcessInstance.SequenceId).Select(x => x.MetaData).FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(metadata))
                {
                    fileName = Common.Utils.ReadJsonTagValue(metadata, Common.Constants.JSON.Tags.Message.Details.Key, Common.Constants.JSON.Tags.Message.Details.ExcelPath);
                    sheetName = fileName.Replace(".xls", "");
                }
            }
        }

        public string GetBranchCode(string reCode)
        {
            string branchCode = string.Empty;
            reCode = Regex.Split(reCode, " ")[0].ToUpper().Trim();
            switch (reCode)
            {
                case "DELHI":
                    branchCode = "108";
                    break;
                case "MUMBAI":
                    branchCode = "111";
                    break;
                case "AHEMDABAD":
                    branchCode = "113";
                    break;
                case "CHANDIGARH":
                    branchCode = "106";
                    break;
            }
            return branchCode;
        }
    }
}
