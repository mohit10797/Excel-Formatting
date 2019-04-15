using Common.Constants;
using DataAccessLayer.DBModel;
using ExcelFormatting.Model;
using NLog;
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

namespace ExcelFormatting
{
    public class ExcelBL
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
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
                foreach (FileInfo file in di.GetFiles(fileName))
                {
                    result = file.FullName.EndsWith(".xls") ? true : false;
                }
            }
            catch (Exception ex)
            { }
            return result;
        }

        /// <summary>
        /// Method to read data from DMS file
        /// </summary>
        /// <returns></returns>
        public List<RCPFromDMSModel> ReadRCPFromDMS(long processInstanceId)
        {
            string metadata = String.Empty, fileName = string.Empty, sheetName = string.Empty;

            using (GPILEntities entities = new GPILEntities())
            {
                metadata = (from pidt in entities.TBL_ProcessInstanceData
                           join pidtl in entities.TBL_ProcessInstanceDetails on pidt.ProcessInstanceId equals pidtl.ProcessInstanceId
                           join psq in entities.TBL_ProcessInstanceDetails on pidt.SequenceId equals psq.SequenceId
                           where pidt.ProcessInstanceId == processInstanceId && pidtl.StateId==Common.Constants.Process.States.RCPExccelProcess.DMSAttachmentDownload
                           orderby pidt.SequenceId descending
                            select pidt.MetaData).ToList().FirstOrDefault();
                //           where 
                //.Where(x => x.ProcessInstanceId == processInstanceId).OrderByDescending(x => x.SequenceId).Select(x => x.MetaData).FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(metadata))
                    fileName = Common.Utils.ReadJsonTagValue(metadata, Common.Constants.JSON.Tags.Message.Details.Key, "FileName");
                sheetName = fileName.Replace(".xls","");
            }
            List<RCPFromDMSModel> rcpFromDMSModelList = new List<RCPFromDMSModel>();
            DataTable data = null;
            try
            {
                bool isOlderExcel = IsOlderVersionExcel(fileName);
                if (isOlderExcel)
                {
                    fileName = this.folderPath + fileName;
                    var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
                    var adapter = new OleDbDataAdapter("SELECT * FROM ["+ sheetName + "$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "FromDMS");
                    data = ds.Tables["FromDMS"];
                    rcpFromDMSModelList = ConvertDataTable<RCPFromDMSModel>(data);
                }
                else
                {
                    fileName = this.folderPath + fileName;
                    FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
                    DataTable dTable = new DataTable();
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets["RCPReport (80)"];
                        int totalRows = workSheet.Dimension.Rows;

                        for (int i = 2; i <= totalRows; i++)
                        {
                            rcpFromDMSModelList.Add(new RCPFromDMSModel
                            {
                                RegionCode = workSheet.Cells[i, 1].Value != null ? workSheet.Cells[i, 1].Value.ToString() : string.Empty,
                                RegionName = workSheet.Cells[i, 2].Value != null ? workSheet.Cells[i, 2].Value.ToString() : string.Empty,
                                TownCode = workSheet.Cells[i, 3].Value != null ? workSheet.Cells[i, 3].Value.ToString() : string.Empty,
                                TownName = workSheet.Cells[i, 4].Value != null ? workSheet.Cells[i, 4].Value.ToString() : string.Empty,
                                Am = workSheet.Cells[i, 5].Value != null ? workSheet.Cells[i, 5].Value.ToString() : string.Empty,
                                DistrCode = workSheet.Cells[i, 6].Value != null ? workSheet.Cells[i, 6].Value.ToString() : string.Empty,
                                RouteName = workSheet.Cells[i, 7].Value != null ? workSheet.Cells[i, 7].Value.ToString() : string.Empty,
                                RouteCode = workSheet.Cells[i, 8].Value != null ? workSheet.Cells[i, 8].Value.ToString() : string.Empty,
                                SalesmanCode = workSheet.Cells[i, 9].Value != null ? workSheet.Cells[i, 9].Value.ToString() : string.Empty,
                                Frequency = workSheet.Cells[i, 10].Value != null ? workSheet.Cells[i, 10].Value.ToString() : string.Empty,
                                Coverage = workSheet.Cells[i, 11].Value != null ? workSheet.Cells[i, 11].Value.ToString() : string.Empty,
                                CompanyOutletCode = workSheet.Cells[i, 12].Value != null ? workSheet.Cells[i, 12].Value.ToString() : string.Empty,
                                CustomerName = workSheet.Cells[i, 13].Value != null ? workSheet.Cells[i, 13].Value.ToString() : string.Empty,
                                CustomerAddress = workSheet.Cells[i, 14].Value != null ? workSheet.Cells[i, 14].Value.ToString() : string.Empty,
                                Category = workSheet.Cells[i, 15].Value != null ? workSheet.Cells[i, 15].Value.ToString() : string.Empty,
                                RetailerType = workSheet.Cells[i, 16].Value != null ? workSheet.Cells[i, 16].Value.ToString() : string.Empty
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            return rcpFromDMSModelList;
        }

        /// <summary>
        /// Method to read data from sales team file
        /// </summary>
        /// <returns></returns>
        public List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam()
        {
            List<RCPFromSalesTeamModel> rcpFromSalesTeamModelList = new List<RCPFromSalesTeamModel>();
            DataTable data = null;
            string fileName = string.Empty;
            try
            {
                bool isOlderExcel = IsOlderVersionExcel("RCP data fromsales team(1)");

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
                    fileName = this.folderPath + "RCP data fromsales team(1).xlsx";
                    FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
                    DataTable dTable = new DataTable();
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets["Sales Propoed Data"];
                        if (workSheet != null && workSheet.Dimension != null)
                        {
                            int totalRows = workSheet.Dimension.Rows;
                            for (int i = 2; i <= totalRows; i++)
                            {
                                rcpFromSalesTeamModelList.Add(new RCPFromSalesTeamModel
                                {
                                    OutletName = workSheet.Cells[i, 1].Value != null ? workSheet.Cells[i, 1].Value.ToString() : string.Empty,
                                    OutletID = workSheet.Cells[i, 2].Value != null ? workSheet.Cells[i, 2].Value.ToString() : string.Empty,
                                    CustomerName = workSheet.Cells[i, 3].Value != null ? workSheet.Cells[i, 3].Value.ToString() : string.Empty,
                                    OutletAddress = workSheet.Cells[i, 4].Value != null ? workSheet.Cells[i, 4].Value.ToString() : string.Empty,
                                    DayofWeek = workSheet.Cells[i, 5].Value != null ? workSheet.Cells[i, 5].Value.ToString() : string.Empty,
                                    NewSEId = workSheet.Cells[i, 6].Value != null ? workSheet.Cells[i, 6].Value.ToString() : string.Empty,
                                    RouteName = workSheet.Cells[i, 7].Value != null ? workSheet.Cells[i, 7].Value.ToString() : string.Empty,
                                    MobileNumber1 = workSheet.Cells[i, 8].Value != null ? workSheet.Cells[i, 8].Value.ToString() : string.Empty,
                                    MobileNumber2 = workSheet.Cells[i, 9].Value != null ? workSheet.Cells[i, 9].Value.ToString() : string.Empty,
                                    ContactPersonName = workSheet.Cells[i, 10].Value != null ? workSheet.Cells[i, 10].Value.ToString() : string.Empty,
                                    Seq = workSheet.Cells[i, 11].Value != null ? workSheet.Cells[i, 11].Value.ToString() : string.Empty,
                                    TSMapping = workSheet.Cells[i, 12].Value != null ? workSheet.Cells[i, 12].Value.ToString() : string.Empty,
                                    OutletClass = workSheet.Cells[i, 13].Value != null ? workSheet.Cells[i, 13].Value.ToString() : string.Empty,
                                    Remark = workSheet.Cells[i, 14].Value != null ? workSheet.Cells[i, 14].Value.ToString() : string.Empty,
                                    Pmo = workSheet.Cells[i, 15].Value != null ? workSheet.Cells[i, 15].Value.ToString() : string.Empty,
                                    DistributedBranchCode = workSheet.Cells[i, 16].Value != null ? workSheet.Cells[i, 16].Value.ToString() : string.Empty,
                                    SECode = workSheet.Cells[i, 17].Value != null ? workSheet.Cells[i, 17].Value.ToString() : string.Empty,
                                    TownCode = workSheet.Cells[i, 18].Value != null ? workSheet.Cells[i, 18].Value.ToString() : string.Empty,
                                    TownName = workSheet.Cells[i, 19].Value != null ? workSheet.Cells[i, 19].Value.ToString() : string.Empty,
                                    SEType = workSheet.Cells[i, 20].Value != null ? workSheet.Cells[i, 20].Value.ToString() : string.Empty,
                                    ReportingToAM = workSheet.Cells[i, 21].Value != null ? workSheet.Cells[i, 21].Value.ToString() : string.Empty,
                                    ReportingLevel = workSheet.Cells[i, 22].Value != null ? workSheet.Cells[i, 22].Value.ToString() : string.Empty,
                                    SalesForceCode = workSheet.Cells[i, 23].Value != null ? workSheet.Cells[i, 23].Value.ToString() : string.Empty,
                                    IsActive = workSheet.Cells[i, 24].Value != null ? workSheet.Cells[i, 24].Value.ToString() : string.Empty,
                                    Category = workSheet.Cells[i, 25].Value != null ? workSheet.Cells[i, 25].Value.ToString() : string.Empty,
                                    JoiningDate = workSheet.Cells[i, 26].Value != null ? workSheet.Cells[i, 26].Value.ToString() : string.Empty,
                                    Email = workSheet.Cells[i, 27].Value != null ? workSheet.Cells[i, 27].Value.ToString() : string.Empty,
                                    RouteCode = workSheet.Cells[i, 28].Value != null ? workSheet.Cells[i, 28].Value.ToString() : string.Empty,
                                    SalesManCode = workSheet.Cells[i, 29].Value != null ? workSheet.Cells[i, 29].Value.ToString() : string.Empty,
                                    RetlrType = workSheet.Cells[i, 30].Value != null ? workSheet.Cells[i, 30].Value.ToString() : string.Empty,
                                    State = workSheet.Cells[i, 31].Value != null ? workSheet.Cells[i, 31].Value.ToString() : string.Empty,
                                    City = workSheet.Cells[i, 32].Value != null ? workSheet.Cells[i, 32].Value.ToString() : string.Empty,
                                    PostalCode = workSheet.Cells[i, 33].Value != null ? workSheet.Cells[i, 33].Value.ToString() : string.Empty,
                                    CustChannelType = workSheet.Cells[i, 34].Value != null ? workSheet.Cells[i, 34].Value.ToString() : string.Empty,
                                    CustChannelSubType = workSheet.Cells[i, 35].Value != null ? workSheet.Cells[i, 35].Value.ToString() : string.Empty,
                                    StoreType = workSheet.Cells[i, 36].Value != null ? workSheet.Cells[i, 36].Value.ToString() : string.Empty,
                                    ProductHierarchyValueCode = workSheet.Cells[i, 37].Value != null ? workSheet.Cells[i, 37].Value.ToString() : string.Empty,
                                    ProductHierarchyCategoryName = workSheet.Cells[i, 38].Value != null ? workSheet.Cells[i, 38].Value.ToString() : string.Empty,
                                    LoginID = workSheet.Cells[i, 39].Value != null ? workSheet.Cells[i, 39].Value.ToString() : string.Empty,
                                    AMCode = workSheet.Cells[i, 40].Value != null ? workSheet.Cells[i, 40].Value.ToString() : string.Empty,
                                    EmpId = workSheet.Cells[i, 41].Value != null ? workSheet.Cells[i, 41].Value.ToString() : string.Empty,
                                    Decision = workSheet.Cells[i, 42].Value != null ? workSheet.Cells[i, 42].Value.ToString() : string.Empty,
                                    RemovalOldId = workSheet.Cells[i, 43].Value != null ? workSheet.Cells[i, 43].Value.ToString() : string.Empty,
                                    RemovalRouteCode = workSheet.Cells[i, 44].Value != null ? workSheet.Cells[i, 44].Value.ToString() : string.Empty,
                                    TransferAccordingToRouteOldId = workSheet.Cells[i, 45].Value != null ? workSheet.Cells[i, 45].Value.ToString() : string.Empty,
                                    TransferAccordingToRouteCode = workSheet.Cells[i, 46].Value != null ? workSheet.Cells[i, 46].Value.ToString() : string.Empty,
                                    OutletIdForRemoval = workSheet.Cells[i, 47].Value != null ? workSheet.Cells[i, 47].Value.ToString() : string.Empty,
                                    OutletIdTransferFromOneSalsemanToAnother = workSheet.Cells[i, 48].Value != null ? workSheet.Cells[i, 48].Value.ToString() : string.Empty,
                                    SalesmanCodeInCaseOfOutletTransferOneToAnotherSalseman = workSheet.Cells[i, 49].Value != null ? workSheet.Cells[i, 49].Value.ToString() : string.Empty,
                                    OutletIdTransferFromOneRouteToAnotherRoute = workSheet.Cells[i, 50].Value != null ? workSheet.Cells[i, 50].Value.ToString() : string.Empty,
                                    RouteCodeInCaseOfOutletTransferFromOneToAnotherRoute = workSheet.Cells[i, 51].Value != null ? workSheet.Cells[i, 51].Value.ToString() : string.Empty,
                                    NewSequence = workSheet.Cells[i, 52].Value != null ? workSheet.Cells[i, 52].Value.ToString() : string.Empty
                                });
                            }
                        }
                        else
                            logger.Info("File from sales team is not exists or is empty.");
                    }
                }

                string fName = "RCP data fromsales team(1).xlsx";
                string sourcePath = @"D:\Shared\RCPExcelFiles";
                string targetPath = @"D:\Shared\Processed";
                //MoveFileToFolder(sourcePath, targetPath, fName);
            }
            catch (Exception ex)
            { }
            return rcpFromSalesTeamModelList;
        }

        private void MoveFileToFolder(string sourcePath, string targetPath, string fileName)
        {
            try
            {
                string date = DateTime.Now.ToString("dd-MM-yyyy") + "\\";
                DirectoryInfo targetFolder = new DirectoryInfo(targetPath);
                DirectoryInfo subFolder = targetFolder.CreateSubdirectory(date);
                if (System.IO.File.Exists(targetPath + date + fileName))
                {
                    System.IO.File.Delete(targetPath + date + fileName);
                }
                File.Move(sourcePath + fileName, targetPath + date + fileName);
            }
            catch (Exception ex)
            { }
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
        public string ApplyingRules(List<RCPFromDMSModel> ReadRCPFromDMS, ref List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, ref List<RCPFromSalesTeamModel> lstFromSalesNewSEId)
        {
            string branchCode = string.Empty;
            string fileRequiredJSON = string.Empty;
            try
            {
                using (GPILEntities entities = new GPILEntities())
                {
                    RCPFromSalesTeamModel tempRcpFromSalesTeamModel = new RCPFromSalesTeamModel();

                    ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(s => !string.IsNullOrWhiteSpace(s.DayofWeek)
                    && !string.IsNullOrWhiteSpace(s.DistributedBranchCode) && !string.IsNullOrWhiteSpace(s.SalesForceCode)
                    && !string.IsNullOrWhiteSpace(s.ReportingToAM) && !string.IsNullOrWhiteSpace(s.Category)
                    && !string.IsNullOrWhiteSpace(s.State) && !string.IsNullOrWhiteSpace(s.City)).ToList();

                    ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(p => ReadRCPFromDMS.Any(x => x.DistrCode == p.DistributedBranchCode)).ToList();

                    var newRecords = ReadRCPFromSalesTeam.Where(x => x.OutletID.ToLower().Contains("new")).ToList();
                    List<string> excelRules = entities.TBL_ExcelRules.Where(x => x.IsActive != null && x.IsActive == true).Select(x => x.RuleName).ToList();

                    //if (excelRules != null && excelRules.Contains("DeleteBlankRows"))
                    //    ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(s => !string.IsNullOrWhiteSpace(s.NewSequence)).ToList();
                    if (excelRules != null && excelRules.Contains("SECode"))
                    {
                        var blankSECode = ReadRCPFromSalesTeam.Where(record => record.SECode == string.Empty).ToList();
                        if (blankSECode != null && blankSECode.Count > 0)
                        {
                            var groupedSECodeFromSalesTeam = blankSECode.GroupBy(x => new { x.DistributedBranchCode }).Select(grp => grp.ToList()).ToList();
                            foreach (var seCode in groupedSECodeFromSalesTeam)
                            {
                                foreach (RCPFromSalesTeamModel fromSales in seCode)
                                {
                                    branchCode = fromSales.Pmo;
                                    CreateSECode(branchCode, blankSECode);
                                }
                            }

                            foreach (RCPFromSalesTeamModel added in blankSECode)
                            {
                                tempRcpFromSalesTeamModel = ReadRCPFromSalesTeam.Where(i => i.SECode == string.Empty).FirstOrDefault();
                                var index = ReadRCPFromSalesTeam.IndexOf(tempRcpFromSalesTeamModel);
                                if (index != -1)
                                    ReadRCPFromSalesTeam[index] = added;
                            }
                        }
                        lstFromSalesNewSEId = blankSECode;
                        fileRequiredJSON = "{\"" + Common.Constants.ExcelFileName.Salesman.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareSalesmanExcel + "\",";

                    }

                    if (excelRules != null && excelRules.Contains("CreateRouteCode"))
                    {
                        int count = 1;
                        string RouteCode = string.Empty;
                        string RouteName = string.Empty;
                        var groupedSECodeFromSalesTeam = ReadRCPFromSalesTeam.GroupBy(x => new { x.SECode, x.DistributedBranchCode, x.DayofWeek }).Select(grp => grp.ToList()).ToList();
                        foreach (var seCode in groupedSECodeFromSalesTeam)
                        {
                            //count = 0001;
                            foreach (RCPFromSalesTeamModel fromSales in seCode)
                            {
                                branchCode = fromSales.Pmo;
                                string bCode = Enum.Parse(typeof(BranchCode), branchCode).ToString();
                                RouteCode = branchCode + DateTime.Now.ToString("ddMMyy") + count.ToString().PadLeft(4, '0');
                                RouteName = bCode + DateTime.Now.ToString("ddMMyy") + count.ToString().PadLeft(4, '0');

                                //fromSales.RouteCode ="";
                                //fromSales.RouteName ="";
                                //CreateRoutCodeAndName(branchCode, ReadRCPFromDMS,ref ReadRCPFromSalesTeam);

                                //tempRcpFromSalesTeamModel = ReadRCPFromSalesTeam.Where(i => i.SECode == string.Empty).FirstOrDefault();
                                var index = ReadRCPFromSalesTeam.IndexOf(fromSales);
                                if (index != -1)
                                {
                                    ReadRCPFromSalesTeam[index].RouteCode = RouteCode;
                                    ReadRCPFromSalesTeam[index].RouteName = RouteName;
                                }
                            }
                            count++;
                        }
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.Beat.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareBeatExcel + "\",";
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.SalesmanRoute.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareSalesmanRouteExcel + "\",";
                    }
                    if (excelRules != null && excelRules.Contains("JoiningDate"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.JoiningDate == string.Empty)
                        .Select(record => { record.JoiningDate = DateTime.Now.ToShortDateString(); return record; }).ToList();
                    }
                    if (excelRules != null && excelRules.Contains("IsActive"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.IsActive == string.Empty)
                        .Select(record => { record.IsActive = "Y"; return record; }).ToList();
                    }
                    if (excelRules != null && excelRules.Contains("Email"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.Email == string.Empty)
                        .Select(record => { record.Email = "abc@abc.com"; return record; }).ToList();
                    }
                    if (excelRules != null && excelRules.Contains("DayOfWeek"))
                    {
                        FormatDayOfWeek(ReadRCPFromSalesTeam);
                    }
                    if ((newRecords != null && newRecords.Count > 0) && excelRules != null && excelRules.Contains("CreateOutletId"))
                    {
                        CreateOutletId(ReadRCPFromDMS, ReadRCPFromSalesTeam);
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.Outlet.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareOutletExcel + "\",";
                    }
                    if (excelRules != null && excelRules.Contains("ReportingLevel"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.ReportingLevel == string.Empty)
                        .Select(record => { record.ReportingLevel = "Assistant Manager"; return record; }).ToList();
                    }
                    if (excelRules != null && excelRules.Contains("MobileNumber"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.MobileNumber1 == string.Empty)
                        .Select(record => { record.MobileNumber1 = "1234567890"; return record; }).ToList();
                    }
                    if (excelRules != null && excelRules.Contains("OutletClass"))
                    {
                        ReadRCPFromSalesTeam.Where(record => record.OutletClass == string.Empty)
                        .Select(record => { record.OutletClass = "M"; return record; }).ToList();
                    }
                    //if (excelRules != null && excelRules.Contains("DistributorBranchCode"))
                    //{
                    //    newRecords.Where(record => record.DistributedBranchCode == string.Empty)
                    //    .Select(record => { record.DistributedBranchCode = ReadRCPFromDMS.FirstOrDefault().DistrCode.Replace("-", ""); return record; }).ToList();
                    //}
                    if (newRecords != null && newRecords.Count > 0)
                    {
                        foreach (RCPFromSalesTeamModel added in newRecords)
                        {
                            tempRcpFromSalesTeamModel = ReadRCPFromSalesTeam.Where(i => i.OutletID == "new").FirstOrDefault();
                            var index = ReadRCPFromSalesTeam.IndexOf(tempRcpFromSalesTeamModel);
                            if (index != -1)
                                ReadRCPFromSalesTeam[index] = added;
                        }
                    }
                    fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.CPCategory.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareCPCategoryExcel + "\",";
                    fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.CustomerRoute.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareCustomerRouteExcel + "\",";
                    fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.BeatPlanning.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareBeatPlanningExcel + "\",";
                    if (ReadRCPFromSalesTeam.Where(x => string.IsNullOrWhiteSpace(x.RemovalOldId)).Count() > 0 ||
                        ReadRCPFromSalesTeam.Where(x => string.IsNullOrWhiteSpace(x.TransferAccordingToRouteOldId)).Count() > 0 ||
                        ReadRCPFromSalesTeam.Where(x => string.IsNullOrWhiteSpace(x.OutletIdTransferFromOneSalsemanToAnother)).Count() > 0)
                    {
                        fileRequiredJSON += "\"" + Common.Constants.ExcelFileName.OutletRemoval.Replace(".xls", "") + "\":\"" + Common.Constants.Process.States.RCPExccelProcess.PrepareORMExcel + "\",";
                    }
                    fileRequiredJSON = fileRequiredJSON.Remove(fileRequiredJSON.Length - 1) + "}";
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
        private void CreateOutletId(List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string tempOutletId = string.Empty;
            int counter = 1;
            try
            {
                foreach (RCPFromSalesTeamModel allFromSalesTeam in ReadRCPFromSalesTeam)
                {
                    if (allFromSalesTeam.OutletID == "new")
                    {
                        string newOutletId = "B0" + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                        while (CheckNewOutletIdExistance(newOutletId, ReadRCPFromSalesTeam) || newOutletId == tempOutletId)
                        {
                            counter++;
                            newOutletId = "B0" + DateTime.Now.ToString("ddMMyy") + "000" + counter;
                            //tempOutletId = newOutletId;
                        }
                        allFromSalesTeam.OutletID = newOutletId;
                        tempOutletId = newOutletId;
                        counter++;
                    }
                }
            }
            catch (Exception ex)
            { }
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
            { }
            //return newRoutCode;
        }

        private void CreateSECode(string branchCode, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string tempSECode = string.Empty, newSECode = string.Empty;
            int counter = 1;
            string bCode = Enum.Parse(typeof(BranchCode), branchCode).ToString();
            try
            {
                foreach (RCPFromSalesTeamModel allFromSalesTeam in ReadRCPFromSalesTeam)
                {
                    newSECode = bCode + "wsm" + DateTime.Now.ToString("ddMMyy") + counter.ToString().PadLeft(3, '0');
                    while (CheckNewSECodeExistance(newSECode, ReadRCPFromSalesTeam) || newSECode == tempSECode)
                    {
                        counter++;
                        newSECode = bCode + "wsm" + DateTime.Now.ToString("ddMMyy") + counter.ToString().PadLeft(3, '0');

                    }
                    allFromSalesTeam.SECode = newSECode;
                    tempSECode = newSECode;
                    counter++;
                    break;
                }
            }
            catch (Exception ex)
            { }
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
                    string[] arrDays = Regex.Split(fromSales.DayofWeek, "-");
                    fromSales.DayofWeek = string.Empty;
                    foreach (string day in arrDays)
                    {
                        if (day.ToLower().Contains("m"))
                            fromSales.DayofWeek += "M";
                        if (day.ToLower().Contains("tu"))
                            fromSales.DayofWeek += "-Tu";
                        if (day.ToLower().Contains("w"))
                            fromSales.DayofWeek += "-W";
                        if (day.ToLower().Contains("th"))
                            fromSales.DayofWeek += "-Th";
                        if (day.ToLower().Contains("f"))
                            fromSales.DayofWeek += "-F";
                        if (day.ToLower().Contains("s"))
                            fromSales.DayofWeek += "-Sa";
                    }
                    fromSales.DayofWeek = fromSales.DayofWeek.StartsWith("-") ? fromSales.DayofWeek.Substring(1) : fromSales.DayofWeek;
                    fromSales.DayofWeek = fromSales.DayofWeek.EndsWith("-") ? fromSales.DayofWeek.Substring(0, fromSales.DayofWeek.Length - 1) : fromSales.DayofWeek;
                }
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// Method to populate excel file after applying the rules
        /// </summary>
        /// <param name="ReadRCPFromSalesTeam"></param>
        public void CreateExcel(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
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
            { }

        }

        /// <summary>
        /// Method to create rows for sales team file
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="ReadRCPFromSalesTeam"></param>
        /// <param name="i"></param>
        private void CreateRawFileRow(ExcelWorksheet ws, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam, int i)
        {
            ws.Cells[i + 2, 1].Value = ReadRCPFromSalesTeam[i].OutletName;
            ws.Cells[i + 2, 2].Value = ReadRCPFromSalesTeam[i].OutletID;
            ws.Cells[i + 2, 3].Value = ReadRCPFromSalesTeam[i].CustomerName;
            ws.Cells[i + 2, 4].Value = ReadRCPFromSalesTeam[i].OutletAddress;
            ws.Cells[i + 2, 5].Value = ReadRCPFromSalesTeam[i].DayofWeek;
            ws.Cells[i + 2, 6].Value = ReadRCPFromSalesTeam[i].NewSEId;
            ws.Cells[i + 2, 7].Value = ReadRCPFromSalesTeam[i].RouteName;
            ws.Cells[i + 2, 8].Value = ReadRCPFromSalesTeam[i].MobileNumber1;
            ws.Cells[i + 2, 9].Value = ReadRCPFromSalesTeam[i].MobileNumber2;
            ws.Cells[i + 2, 10].Value = ReadRCPFromSalesTeam[i].ContactPersonName;
            ws.Cells[i + 2, 11].Value = ReadRCPFromSalesTeam[i].Seq;
            ws.Cells[i + 2, 12].Value = ReadRCPFromSalesTeam[i].TSMapping;
            ws.Cells[i + 2, 13].Value = ReadRCPFromSalesTeam[i].OutletClass;
            ws.Cells[i + 2, 14].Value = ReadRCPFromSalesTeam[i].Remark;
            ws.Cells[i + 2, 15].Value = ReadRCPFromSalesTeam[i].Pmo;
            ws.Cells[i + 2, 16].Value = ReadRCPFromSalesTeam[i].DistributedBranchCode;
            ws.Cells[i + 2, 17].Value = ReadRCPFromSalesTeam[i].SECode;
            ws.Cells[i + 2, 18].Value = ReadRCPFromSalesTeam[i].TownCode;
            ws.Cells[i + 2, 19].Value = ReadRCPFromSalesTeam[i].TownName;
            ws.Cells[i + 2, 20].Value = ReadRCPFromSalesTeam[i].SEType;
            ws.Cells[i + 2, 21].Value = ReadRCPFromSalesTeam[i].ReportingToAM;
            ws.Cells[i + 2, 22].Value = ReadRCPFromSalesTeam[i].ReportingLevel;
            ws.Cells[i + 2, 23].Value = ReadRCPFromSalesTeam[i].SalesForceCode;
            ws.Cells[i + 2, 24].Value = ReadRCPFromSalesTeam[i].IsActive;
            ws.Cells[i + 2, 25].Value = ReadRCPFromSalesTeam[i].Category;
            ws.Cells[i + 2, 26].Value = ReadRCPFromSalesTeam[i].JoiningDate;
            ws.Cells[i + 2, 27].Value = ReadRCPFromSalesTeam[i].Email;
            ws.Cells[i + 2, 28].Value = ReadRCPFromSalesTeam[i].RouteCode;
            ws.Cells[i + 2, 29].Value = ReadRCPFromSalesTeam[i].SalesManCode;
            ws.Cells[i + 2, 30].Value = ReadRCPFromSalesTeam[i].RetlrType;
            ws.Cells[i + 2, 31].Value = ReadRCPFromSalesTeam[i].State;
            ws.Cells[i + 2, 32].Value = ReadRCPFromSalesTeam[i].City;
            ws.Cells[i + 2, 33].Value = ReadRCPFromSalesTeam[i].PostalCode;
            ws.Cells[i + 2, 34].Value = ReadRCPFromSalesTeam[i].CustChannelType;
            ws.Cells[i + 2, 35].Value = ReadRCPFromSalesTeam[i].CustChannelSubType;
            ws.Cells[i + 2, 36].Value = ReadRCPFromSalesTeam[i].StoreType;
            ws.Cells[i + 2, 37].Value = ReadRCPFromSalesTeam[i].ProductHierarchyValueCode;
            ws.Cells[i + 2, 38].Value = ReadRCPFromSalesTeam[i].ProductHierarchyCategoryName;
            ws.Cells[i + 2, 39].Value = ReadRCPFromSalesTeam[i].LoginID;
            ws.Cells[i + 2, 40].Value = ReadRCPFromSalesTeam[i].AMCode;
            ws.Cells[i + 2, 41].Value = ReadRCPFromSalesTeam[i].EmpId;
            ws.Cells[i + 2, 42].Value = ReadRCPFromSalesTeam[i].Decision;
            ws.Cells[i + 2, 43].Value = ReadRCPFromSalesTeam[i].RemovalOldId;
            ws.Cells[i + 2, 44].Value = ReadRCPFromSalesTeam[i].RemovalRouteCode;
            ws.Cells[i + 2, 45].Value = ReadRCPFromSalesTeam[i].TransferAccordingToRouteOldId;
            ws.Cells[i + 2, 46].Value = ReadRCPFromSalesTeam[i].TransferAccordingToRouteCode;
            ws.Cells[i + 2, 47].Value = ReadRCPFromSalesTeam[i].OutletIdForRemoval;
            ws.Cells[i + 2, 48].Value = ReadRCPFromSalesTeam[i].OutletIdTransferFromOneSalsemanToAnother;
            ws.Cells[i + 2, 49].Value = ReadRCPFromSalesTeam[i].SalesmanCodeInCaseOfOutletTransferOneToAnotherSalseman;
            ws.Cells[i + 2, 50].Value = ReadRCPFromSalesTeam[i].OutletIdTransferFromOneRouteToAnotherRoute;
            ws.Cells[i + 2, 51].Value = ReadRCPFromSalesTeam[i].RouteCodeInCaseOfOutletTransferFromOneToAnotherRoute;
            ws.Cells[i + 2, 52].Value = ReadRCPFromSalesTeam[i].NewSequence;
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
        }

        private void CreateSalsemanRouteRawFileRow(ExcelWorksheet ws, List<SalesmanRouteModel> salsemanRouteModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = salsemanRouteModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = salsemanRouteModelList[i].SalesmanCode;
            ws.Cells[i + 2, 3].Value = salsemanRouteModelList[i].RouteCode;
        }

        private void CreateOutletRouteRemovalRawFileRow(ExcelWorksheet ws, List<OutletRouteRemovalModel> outletRouteRemovalModelList, int i)
        {
            ws.Cells[i + 2, 1].Value = outletRouteRemovalModelList[i].DistributorBranchCode;
            ws.Cells[i + 2, 2].Value = outletRouteRemovalModelList[i].OutletCode;
            ws.Cells[i + 2, 3].Value = outletRouteRemovalModelList[i].RouteCode;
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
            ws.Cells[1, 1].Value = "Outlet Name";
            ws.Cells[1, 2].Value = "Outlet ID";
            ws.Cells[1, 3].Value = "Customer Name";
            ws.Cells[1, 4].Value = "Outlet Address";
            ws.Cells[1, 5].Value = "Day of Week";
            ws.Cells[1, 6].Value = "New SE Id";
            ws.Cells[1, 7].Value = "Route Name";
            ws.Cells[1, 8].Value = "Mobile Number1";
            ws.Cells[1, 9].Value = "Mobile Number2";
            ws.Cells[1, 10].Value = "Contact Person Name";
            ws.Cells[1, 11].Value = "Seq";
            ws.Cells[1, 12].Value = "TSMapping";
            ws.Cells[1, 13].Value = "Outlet Class";
            ws.Cells[1, 14].Value = "Remark";
            ws.Cells[1, 15].Value = "Pmo";
            ws.Cells[1, 16].Value = "Distributed Branch Code";
            ws.Cells[1, 17].Value = "SE Code";
            ws.Cells[1, 18].Value = "Town Code";
            ws.Cells[1, 19].Value = "Town Name";
            ws.Cells[1, 20].Value = "SE Type";
            ws.Cells[1, 21].Value = "Reporting To AM";
            ws.Cells[1, 22].Value = "Reporting Level";
            ws.Cells[1, 23].Value = "SalesForce Code";
            ws.Cells[1, 24].Value = "IsActive";
            ws.Cells[1, 25].Value = "Category";
            ws.Cells[1, 26].Value = "Joining Date";
            ws.Cells[1, 27].Value = "Email";
            ws.Cells[1, 28].Value = "Route Code";
            ws.Cells[1, 29].Value = "SalesMan Code";
            ws.Cells[1, 30].Value = "Retlr Type";
            ws.Cells[1, 31].Value = "State";
            ws.Cells[1, 32].Value = "City";
            ws.Cells[1, 33].Value = "Postal Code";
            ws.Cells[1, 34].Value = "Cust Channel Type";
            ws.Cells[1, 35].Value = "Cust Channel Sub Type";
            ws.Cells[1, 36].Value = "Store Type";
            ws.Cells[1, 37].Value = "Product Hierarchy Value Code";
            ws.Cells[1, 38].Value = "Product Hierarchy Category Name";
            ws.Cells[1, 39].Value = "Login ID";
            ws.Cells[1, 40].Value = "AM Code";
            ws.Cells[1, 41].Value = "Emp Id";
            ws.Cells[1, 42].Value = "Decision";
            ws.Cells[1, 43].Value = "Removal Old Id";
            ws.Cells[1, 44].Value = "Removal Route Code";
            ws.Cells[1, 45].Value = "Transfer According To Route Old Id";
            ws.Cells[1, 46].Value = "Transfer According To Route Code";
            ws.Cells[1, 47].Value = "Outlet Id For Removal";
            ws.Cells[1, 48].Value = "Outlet Id Transfer From One Salseman To Another";
            ws.Cells[1, 49].Value = "Salesman Code In Case Of Outlet Transfer One To Another Salseman";
            ws.Cells[1, 50].Value = "Outlet Id Transfer From One Route To Another Route";
            ws.Cells[1, 51].Value = "Route Code In Case Of Outlet Transfer From One To Another Route";
            ws.Cells[1, 52].Value = "New Sequence";
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
                          where fromSalesTeam.OutletID == newOutletId
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
        }

        private bool CheckNewSECodeExistance(string newSEId, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            var exists = (from fromSalesTeam in ReadRCPFromSalesTeam
                              //join fromSalesTeam in ReadRCPFromSalesTeam on fromDMS.DistrCode equals fromSalesTeam.DistributedBranchCode
                          where fromSalesTeam.SECode == newSEId //|| fromDMS.RouteCode == newRoutCode
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
            //return readRCPFromSalesTeam.SECode != newSEId;
        }

        private bool CheckNewRouteCodeExistance(string newRoutCode, List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            var exists = (from fromSalesTeam in ReadRCPFromSalesTeam
                              //join fromSalesTeam in ReadRCPFromSalesTeam on fromDMS.DistrCode equals fromSalesTeam.DistributedBranchCode
                          where fromSalesTeam.RouteCode == newRoutCode //|| fromDMS.RouteCode == newRoutCode
                          select fromSalesTeam).ToList();

            return exists != null && exists.Count > 0;
        }

        public List<RCPFromSalesTeamModel> ReadRCPFromFormattedSalesFile()
        {
            string fileName = string.Empty;
            List<RCPFromSalesTeamModel> rcpFromSalesTeamModelList = new List<RCPFromSalesTeamModel>();
            fileName = this.folderPath + "RCP data fromsales team(1)Updated.xlsx";
            FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
            DataTable dTable = new DataTable();
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();//["Sales Propoed Data"];
                    int totalRows = workSheet.Dimension.Rows;

                    for (int i = 2; i <= totalRows; i++)
                    {
                        rcpFromSalesTeamModelList.Add(new RCPFromSalesTeamModel
                        {
                            OutletName = workSheet.Cells[i, 1].Value != null ? workSheet.Cells[i, 1].Value.ToString() : string.Empty,
                            OutletID = workSheet.Cells[i, 2].Value != null ? workSheet.Cells[i, 2].Value.ToString() : string.Empty,
                            CustomerName = workSheet.Cells[i, 3].Value != null ? workSheet.Cells[i, 3].Value.ToString() : string.Empty,
                            OutletAddress = workSheet.Cells[i, 4].Value != null ? workSheet.Cells[i, 4].Value.ToString() : string.Empty,
                            DayofWeek = workSheet.Cells[i, 5].Value != null ? workSheet.Cells[i, 5].Value.ToString() : string.Empty,
                            NewSEId = workSheet.Cells[i, 6].Value != null ? workSheet.Cells[i, 6].Value.ToString() : string.Empty,
                            RouteName = workSheet.Cells[i, 7].Value != null ? workSheet.Cells[i, 7].Value.ToString() : string.Empty,
                            MobileNumber1 = workSheet.Cells[i, 8].Value != null ? workSheet.Cells[i, 8].Value.ToString() : string.Empty,
                            MobileNumber2 = workSheet.Cells[i, 9].Value != null ? workSheet.Cells[i, 9].Value.ToString() : string.Empty,
                            ContactPersonName = workSheet.Cells[i, 10].Value != null ? workSheet.Cells[i, 10].Value.ToString() : string.Empty,
                            Seq = workSheet.Cells[i, 11].Value != null ? workSheet.Cells[i, 11].Value.ToString() : string.Empty,
                            TSMapping = workSheet.Cells[i, 12].Value != null ? workSheet.Cells[i, 12].Value.ToString() : string.Empty,
                            OutletClass = workSheet.Cells[i, 13].Value != null ? workSheet.Cells[i, 13].Value.ToString() : string.Empty,
                            Remark = workSheet.Cells[i, 14].Value != null ? workSheet.Cells[i, 14].Value.ToString() : string.Empty,
                            Pmo = workSheet.Cells[i, 15].Value != null ? workSheet.Cells[i, 15].Value.ToString() : string.Empty,
                            DistributedBranchCode = workSheet.Cells[i, 16].Value != null ? workSheet.Cells[i, 16].Value.ToString() : string.Empty,
                            SECode = workSheet.Cells[i, 17].Value != null ? workSheet.Cells[i, 17].Value.ToString() : string.Empty,
                            TownCode = workSheet.Cells[i, 18].Value != null ? workSheet.Cells[i, 18].Value.ToString() : string.Empty,
                            TownName = workSheet.Cells[i, 19].Value != null ? workSheet.Cells[i, 19].Value.ToString() : string.Empty,
                            SEType = workSheet.Cells[i, 20].Value != null ? workSheet.Cells[i, 20].Value.ToString() : string.Empty,
                            ReportingToAM = workSheet.Cells[i, 21].Value != null ? workSheet.Cells[i, 21].Value.ToString() : string.Empty,
                            ReportingLevel = workSheet.Cells[i, 22].Value != null ? workSheet.Cells[i, 22].Value.ToString() : string.Empty,
                            SalesForceCode = workSheet.Cells[i, 23].Value != null ? workSheet.Cells[i, 23].Value.ToString() : string.Empty,
                            IsActive = workSheet.Cells[i, 24].Value != null ? workSheet.Cells[i, 24].Value.ToString() : string.Empty,
                            Category = workSheet.Cells[i, 25].Value != null ? workSheet.Cells[i, 25].Value.ToString() : string.Empty,
                            JoiningDate = workSheet.Cells[i, 26].Value != null ? workSheet.Cells[i, 26].Value.ToString() : string.Empty,
                            Email = workSheet.Cells[i, 27].Value != null ? workSheet.Cells[i, 27].Value.ToString() : string.Empty,
                            RouteCode = workSheet.Cells[i, 28].Value != null ? workSheet.Cells[i, 28].Value.ToString() : string.Empty,
                            SalesManCode = workSheet.Cells[i, 29].Value != null ? workSheet.Cells[i, 29].Value.ToString() : string.Empty,
                            RetlrType = workSheet.Cells[i, 30].Value != null ? workSheet.Cells[i, 30].Value.ToString() : string.Empty,
                            State = workSheet.Cells[i, 31].Value != null ? workSheet.Cells[i, 31].Value.ToString() : string.Empty,
                            City = workSheet.Cells[i, 32].Value != null ? workSheet.Cells[i, 32].Value.ToString() : string.Empty,
                            PostalCode = workSheet.Cells[i, 33].Value != null ? workSheet.Cells[i, 33].Value.ToString() : string.Empty,
                            CustChannelType = workSheet.Cells[i, 34].Value != null ? workSheet.Cells[i, 34].Value.ToString() : string.Empty,
                            CustChannelSubType = workSheet.Cells[i, 35].Value != null ? workSheet.Cells[i, 35].Value.ToString() : string.Empty,
                            StoreType = workSheet.Cells[i, 36].Value != null ? workSheet.Cells[i, 36].Value.ToString() : string.Empty,
                            ProductHierarchyValueCode = workSheet.Cells[i, 37].Value != null ? workSheet.Cells[i, 37].Value.ToString() : string.Empty,
                            ProductHierarchyCategoryName = workSheet.Cells[i, 38].Value != null ? workSheet.Cells[i, 38].Value.ToString() : string.Empty,
                            LoginID = workSheet.Cells[i, 39].Value != null ? workSheet.Cells[i, 39].Value.ToString() : string.Empty,
                            AMCode = workSheet.Cells[i, 40].Value != null ? workSheet.Cells[i, 40].Value.ToString() : string.Empty,
                            EmpId = workSheet.Cells[i, 41].Value != null ? workSheet.Cells[i, 41].Value.ToString() : string.Empty,
                            Decision = workSheet.Cells[i, 42].Value != null ? workSheet.Cells[i, 42].Value.ToString() : string.Empty,
                            RemovalOldId = workSheet.Cells[i, 43].Value != null ? workSheet.Cells[i, 43].Value.ToString() : string.Empty,
                            RemovalRouteCode = workSheet.Cells[i, 44].Value != null ? workSheet.Cells[i, 44].Value.ToString() : string.Empty,
                            TransferAccordingToRouteOldId = workSheet.Cells[i, 45].Value != null ? workSheet.Cells[i, 45].Value.ToString() : string.Empty,
                            TransferAccordingToRouteCode = workSheet.Cells[i, 46].Value != null ? workSheet.Cells[i, 46].Value.ToString() : string.Empty,
                            OutletIdForRemoval = workSheet.Cells[i, 47].Value != null ? workSheet.Cells[i, 47].Value.ToString() : string.Empty,
                            OutletIdTransferFromOneSalsemanToAnother = workSheet.Cells[i, 48].Value != null ? workSheet.Cells[i, 48].Value.ToString() : string.Empty,
                            SalesmanCodeInCaseOfOutletTransferOneToAnotherSalseman = workSheet.Cells[i, 49].Value != null ? workSheet.Cells[i, 49].Value.ToString() : string.Empty,
                            OutletIdTransferFromOneRouteToAnotherRoute = workSheet.Cells[i, 50].Value != null ? workSheet.Cells[i, 50].Value.ToString() : string.Empty,
                            RouteCodeInCaseOfOutletTransferFromOneToAnotherRoute = workSheet.Cells[i, 51].Value != null ? workSheet.Cells[i, 51].Value.ToString() : string.Empty,
                            NewSequence = workSheet.Cells[i, 52].Value != null ? workSheet.Cells[i, 52].Value.ToString() : string.Empty
                        });
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return rcpFromSalesTeamModelList;
        }

        public void CreateSalesman(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                List<SalesmanModel> salesmanModelList = new List<SalesmanModel>();
                bool isOlderExcel = IsOlderVersionExcel("Salesman");
                if (isOlderExcel)
                {
                    FillSalesman(salesmanModelList, ReadRCPFromSalesTeam);
                    CreateSalesmanFileHeader(ws);
                    for (int i = 0; i < salesmanModelList.Count; i++)
                    {
                        CreateSalesmanRawFileRow(ws, salesmanModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.Salesman));
                }
                else
                {
                    CreateSalesmanFile(ReadRCPFromSalesTeam);
                }
            }
            catch (Exception ex)
            { }
            string fName = "Salesman.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public void CreateCustomerProductCategory(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<CustomerProductCategoryModel> customerProductCategoryModelList = new List<CustomerProductCategoryModel>();
                bool isOlderExcel = IsOlderVersionExcel("CustomerProductCategory");
                if (isOlderExcel)
                {
                    FillCustomerProductCategory(customerProductCategoryModelList, ReadRCPFromSalesTeam);
                    CreateCustomerProductCategoryFileHeader(ws);
                    for (int i = 0; i < customerProductCategoryModelList.Count; i++)
                    {
                        CreateCustomerProductCategoryRawFileRow(ws, customerProductCategoryModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.CPCategory));
                }
                else
                {
                    CreateCustomerProductCategoryFile(ReadRCPFromSalesTeam);
                }
            }
            catch (Exception ex)
            { }
            string fName = "CustomerProductCategory.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public void CreateOutletRemoval(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<OutletRouteRemovalModel> outletRouteRemovalModellList = new List<OutletRouteRemovalModel>();
                bool isOlderExcel = IsOlderVersionExcel("OutletRouteRemoval");
                if (isOlderExcel)
                {
                    ReadRCPFromSalesTeam = ReadRCPFromSalesTeam.Where(record => record.RemovalOldId != string.Empty).ToList();
                    outletRouteRemovalModellList = ReadRCPFromSalesTeam.Select(record => new OutletRouteRemovalModel
                    {
                        DistributorBranchCode = record.DistributedBranchCode,
                        OutletCode = record.RemovalOldId,
                        RouteCode = record.RemovalRouteCode
                    }).ToList();
                    CreateOutletRemovalFileHeader(ws);
                    for (int i = 0; i < outletRouteRemovalModellList.Count; i++)
                    {
                        CreateOutletRouteRemovalRawFileRow(ws, outletRouteRemovalModellList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.OutletRemoval));
                }
                else
                {
                    //CreateCustomerProductCategoryFile(ReadRCPFromSalesTeam);
                }
            }
            catch (Exception ex)
            { }
            string fName = "CustomerProductCategory.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public void CreateBeat(List<RCPFromDMSModel> ReadRCPFromDMS, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<BeatModel> beatModelList = new List<BeatModel>();
                bool isOlderExcel = IsOlderVersionExcel("Beat");
                if (isOlderExcel)
                {
                    FillBeat(beatModelList, ReadRCPFromDMS, ReadRCPFromSalesTeam);
                    CreateBeatFileHeader(ws);

                    for (int i = 0; i < beatModelList.Count; i++)
                    {
                        CreateBeatRawFileRow(ws, beatModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.Beat));

                }
                else
                {
                    CreateBeatFile(ReadRCPFromSalesTeam);
                }
                string fName = "Beat.xls";
                string sourcePath = @"D:\Shared\RCPExcelFiles\";
                string targetPath = @"D:\Shared\Processed\";
                //MoveFileToFolder(sourcePath, targetPath, fName);
            }
            catch (Exception ex)
            { }

        }

        public void CreateSalesmanRoute(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<SalesmanRouteModel> salesmanRouteModelList = new List<SalesmanRouteModel>();
                bool isOlderExcel = IsOlderVersionExcel("SalesmanRoute");
                if (isOlderExcel)
                {
                    FillSalsemanRoute(salesmanRouteModelList, ReadRCPFromSalesTeam);
                    CreateSalsemanRouteFileHeader(ws);
                    for (int i = 0; i < salesmanRouteModelList.Count; i++)
                    {
                        CreateSalsemanRouteRawFileRow(ws, salesmanRouteModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.SalesmanRoute));
                }
                else
                {
                    CreateSalesmanRouteFile(ReadRCPFromSalesTeam);
                }
            }
            catch (Exception ex)
            { }
            string fName = "SalesmanRoute.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public void CreateOutletRouteRemoval(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<OutletRouteRemovalModel> outletRouteRemovalModelList = new List<OutletRouteRemovalModel>();
                bool isOlderExcel = IsOlderVersionExcel("OutletRouteRemoval");
                if (isOlderExcel)
                {
                    FillOutletRouteRemoval(outletRouteRemovalModelList, ReadRCPFromSalesTeam);
                    CreateOutletRemovalFileHeader(ws);
                    for (int i = 0; i < outletRouteRemovalModelList.Count; i++)
                    {
                        CreateOutletRouteRemovalRawFileRow(ws, outletRouteRemovalModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.OutletRemoval));
                }
            }
            catch (Exception ex)
            { }
            string fName = "SalesmanRoute.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public void CreateCustomerRoute(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            List<CustomerRouteModel> customerRouteModelList = new List<CustomerRouteModel>();
            List<CustomerRouteModel> tempcustomerRouteModelList = new List<CustomerRouteModel>();
            CustomerRouteModel tempRcpFromSalesTeamModel = null;
            try
            {
                customerRouteModelList = ReadCustomerRoutefromDMS();
                int MaxCoverSeq = 0;
                List<string> DistCode = ReadRCPFromSalesTeam.Select(x => x.DistributedBranchCode).Distinct().ToList();
                foreach (string item in DistCode)
                {
                    MaxCoverSeq = Convert.ToInt32(customerRouteModelList.Where(x => x.DistributorCode.Equals(item)).
                        OrderByDescending(x => x.MaximumCoverageSequenceNo).Select(x => x.MaximumCoverageSequenceNo).FirstOrDefault());
                    foreach (var fromSales in ReadRCPFromSalesTeam)
                    {
                        tempRcpFromSalesTeamModel = new CustomerRouteModel();
                        if (fromSales.DistributedBranchCode.Equals(item))
                        {
                            tempRcpFromSalesTeamModel.DistributorBranchCode = fromSales.DistributedBranchCode;
                            tempRcpFromSalesTeamModel.OutletCode = fromSales.OutletID;
                            tempRcpFromSalesTeamModel.RouteCode = fromSales.RouteCode;
                            tempRcpFromSalesTeamModel.CoverageSequence = (++MaxCoverSeq).ToString();
                            tempcustomerRouteModelList.Add(tempRcpFromSalesTeamModel);
                        }
                        else { }
                    }
                }

                bool isOlderExcel = IsOlderVersionExcel("CustomerRoute");
                if (isOlderExcel)
                {
                    ExcelPackage ExcelPkg = new ExcelPackage();
                    ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                    ExcelWorksheet ws1 = ExcelPkg.Workbook.Worksheets.Add("Details sheet");
                    CreateCustomerRouteFileHeader(ws, ws1);
                    for (int i = 0; i < tempcustomerRouteModelList.Count; i++)
                    {
                        CreateCustomerRouteRawFileRow(ws, tempcustomerRouteModelList, i);
                    }
                    for (int i = 0; i < customerRouteModelList.Count; i++)
                    {
                        CreateCustomerRouteRawFileRow1(ws1, customerRouteModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ws1.Protection.IsProtected = false;
                    ws1.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.CustomerRoute));
                }
            }
            catch (Exception ex)
            { }
            string fName = "CustomerRoute.xls";
            string sourcePath = @"D:\Shared\RCPExcelFiles\";
            string targetPath = @"D:\Shared\Processed\";
            //MoveFileToFolder(sourcePath, targetPath, fName);
        }

        public List<BeatPlanningModel> CreateBeatPlanning(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            string json = string.Empty;
            List<BeatPlanningModel> beatPlanningModelList = new List<BeatPlanningModel>();
            BeatPlanningModel beatPlanningModel = null;
            try
            {
                var tempRCPFromSalesTeam = (from c in ReadRCPFromSalesTeam
                                            group c by new { c.DistributedBranchCode, c.DayofWeek, c.RouteName, c.SECode, c.AMCode, c.RouteCode } into grp
                                            select new
                                            {
                                                grp.Key.DistributedBranchCode,
                                                grp.Key.DayofWeek,
                                                grp.Key.RouteName,
                                                grp.Key.SECode,
                                                grp.Key.AMCode,
                                                grp.Key.RouteCode,


                                            }).Distinct().ToList();

                for (int i = 0; i < tempRCPFromSalesTeam.Count; i++)
                {
                    beatPlanningModel = new BeatPlanningModel();
                    beatPlanningModel.SECode = tempRCPFromSalesTeam[i].SECode;
                    beatPlanningModel.WDCode = tempRCPFromSalesTeam[i].DistributedBranchCode;
                    beatPlanningModel.RouteName = tempRCPFromSalesTeam[i].RouteName;
                    beatPlanningModel.VisitDates = GetVisitDatesFromDaysOfweek(tempRCPFromSalesTeam[i].DayofWeek);
                    beatPlanningModel.EndDate = GetEndDateFromDaysOfweek(tempRCPFromSalesTeam[i].DayofWeek);
                    beatPlanningModelList.Add(beatPlanningModel);
                }

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
                    ws.Cells[i + 2, 2].Value = tempRCPFromSalesTeam[i].DayofWeek;
                    ws.Cells[i + 2, 3].Value = tempRCPFromSalesTeam[i].DistributedBranchCode;
                    ws.Cells[i + 2, 4].Value = tempRCPFromSalesTeam[i].AMCode;
                    ws.Cells[i + 2, 5].Value = tempRCPFromSalesTeam[i].RouteCode;
                    ws.Cells[i + 2, 6].Value = tempRCPFromSalesTeam[i].RouteName;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + Common.Constants.ExcelFileName.BeatPlanning));
            }
            catch (Exception ex)
            {

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
                case "S": result = 6; break;
            }
            return result;
        }

        private void FillSalesman(List<SalesmanModel> SalesmanModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            SalesmanModel salesmanModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    salesmanModel = new SalesmanModel();
                    salesmanModel.DistributorBranchCode = ReadRCPFromSalesTeam[i].DistributedBranchCode;
                    salesmanModel.SECode = ReadRCPFromSalesTeam[i].SECode;
                    salesmanModel.SalesmanCode = ReadRCPFromSalesTeam[i].SECode;
                    salesmanModel.SalesmanName = ReadRCPFromSalesTeam[i].SECode;
                    salesmanModel.SEType = ReadRCPFromSalesTeam[i].SEType;
                    salesmanModel.ReportingTo = ReadRCPFromSalesTeam[i].ReportingToAM;
                    salesmanModel.ReportingLevel = ReadRCPFromSalesTeam[i].ReportingLevel;
                    salesmanModel.SalesForceCode = ReadRCPFromSalesTeam[i].SalesForceCode;
                    salesmanModel.IsActive = ReadRCPFromSalesTeam[i].IsActive;
                    salesmanModel.Category = ReadRCPFromSalesTeam[i].Category;
                    salesmanModel.JoiningDate = ReadRCPFromSalesTeam[i].JoiningDate;
                    salesmanModel.Email = ReadRCPFromSalesTeam[i].Email;
                    salesmanModel.PhoneNo = ReadRCPFromSalesTeam[i].MobileNumber1;
                    SalesmanModelList.Add(salesmanModel);
                }
            }
            catch (Exception ex)
            { }
        }

        private void FillOutlet(List<OutletModel> outletModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            OutletModel outletModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    outletModel = new OutletModel();
                    outletModel.CashDiscPerc = "0";
                    outletModel.City = ReadRCPFromSalesTeam[i].City;
                    outletModel.Country = "India";
                    outletModel.CreditBillAct = "None";
                    outletModel.CreditBills = "None";
                    outletModel.CreditDaysAct = "None";
                    outletModel.CreditLimitAct = "None";
                    outletModel.CustChannelSubType = ReadRCPFromSalesTeam[i].CustChannelSubType;
                    outletModel.CustChannelType = ReadRCPFromSalesTeam[i].CustChannelType;
                    outletModel.DistributorBranchCode = ReadRCPFromSalesTeam[i].DistributedBranchCode;
                    outletModel.Email = ReadRCPFromSalesTeam[i].Email;
                    outletModel.EnrollDate = string.Empty;
                    outletModel.IsActive = ReadRCPFromSalesTeam[i].IsActive;
                    outletModel.OutletAddress1 = ReadRCPFromSalesTeam[i].OutletAddress;
                    outletModel.OutletAddress2 = string.Empty;
                    outletModel.OutletAddress3 = string.Empty;
                    outletModel.OutletCode = ReadRCPFromSalesTeam[i].OutletID;
                    outletModel.OutletName = ReadRCPFromSalesTeam[i].OutletName;
                    outletModel.PhoneNo = ReadRCPFromSalesTeam[i].MobileNumber1;
                    outletModel.PostalCode = ReadRCPFromSalesTeam[i].PostalCode;
                    outletModel.RetlrType = string.IsNullOrWhiteSpace(ReadRCPFromSalesTeam[i].RetlrType) ? "Retailer" : ReadRCPFromSalesTeam[i].RetlrType;
                    outletModel.State = ReadRCPFromSalesTeam[i].State;
                    outletModel.StoreType = ReadRCPFromSalesTeam[i].StoreType;
                    outletModel.TaxType = "VAT";
                    outletModelList.Add(outletModel);
                }
            }
            catch (Exception ex)
            { }
        }

        private void FillCustomerProductCategory(List<CustomerProductCategoryModel> customerProductCategoryModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            CustomerProductCategoryModel customerProductCategoryModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    customerProductCategoryModel = new CustomerProductCategoryModel();
                    customerProductCategoryModel.OutletCode = ReadRCPFromSalesTeam[i].OutletName;
                    customerProductCategoryModel.ProductHierarchyCategoryName = ReadRCPFromSalesTeam[i].ProductHierarchyCategoryName;
                    customerProductCategoryModel.ProductHierarchyLevelCode = "100";
                    customerProductCategoryModel.ProductHierarchyValueCode = ReadRCPFromSalesTeam[i].ProductHierarchyValueCode;
                    customerProductCategoryModelList.Add(customerProductCategoryModel);
                }
            }
            catch (Exception ex)
            { }
        }

        private void FillOutletRemoval(List<OutletRouteRemovalModel> outletRemovalModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            OutletRouteRemovalModel outletRemovalModel = null;
            try
            {
                for (int i = 0; i < ReadRCPFromSalesTeam.Count; i++)
                {
                    outletRemovalModel = new OutletRouteRemovalModel();
                    outletRemovalModel.OutletCode = ReadRCPFromSalesTeam[i].OutletID;
                    outletRemovalModel.RouteCode = ReadRCPFromSalesTeam[i].RouteCode;
                    outletRemovalModel.DistributorBranchCode = ReadRCPFromSalesTeam[i].DistributedBranchCode;
                    outletRemovalModelList.Add(outletRemovalModel);
                }
            }
            catch (Exception ex)
            { }
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
                    beatModel.Distance = "50";
                    beatModel.Population = "50";
                    BeatModelList.Add(beatModel);
                }

            }
            catch (Exception ex)
            { }
        }

        private void FillSalsemanRoute(List<SalesmanRouteModel> salesmanRouteModelList, List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            SalesmanRouteModel salesmanRouteModel = null;
            try
            {
                var tempsalesmanRouteModelList = ReadRCPFromSalesTeam.GroupBy(x => new
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
            { }
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
                    outletRouteRemovalModel.OutletCode = ReadRCPFromSalesTeam[i].OutletID;
                    outletRouteRemovalModel.RouteCode = ReadRCPFromSalesTeam[i].RouteCode;
                    outletRouteRemovalModelList.Add(outletRouteRemovalModel);
                }
            }
            catch (Exception ex)
            { }
        }

        public List<CustomerRouteModel> ReadCustomerRoutefromDMS()
        {
            List<CustomerRouteModel> cutomerRoutefromDMS = new List<CustomerRouteModel>();
            DataTable data = null;
            string fileName = string.Empty;
            try
            {
                bool isOlderExcel = IsOlderVersionExcel("CustomerRoute");
                if (isOlderExcel)
                {
                    fileName = this.folderPath + "CustomerRoute.xls";
                    var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", fileName);
                    var adapter = new OleDbDataAdapter("SELECT * FROM [Details sheet$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "FromDMS");
                    data = ds.Tables["FromDMS"];
                    cutomerRoutefromDMS = ConvertDataTable<CustomerRouteModel>(data);
                }
                else
                {
                    fileName = this.folderPath + "CustomerRoute.xlsx";
                    FileInfo file = new FileInfo(Path.Combine(this.folderPath, fileName));
                    DataTable dTable = new DataTable();
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets["Details sheet"];
                        int totalRows = workSheet.Dimension.Rows;

                        for (int i = 2; i <= totalRows; i++)
                        {
                            cutomerRoutefromDMS.Add(new CustomerRouteModel
                            {
                                DistributorCode = workSheet.Cells[i, 1].Value != null ? workSheet.Cells[i, 1].Value.ToString() : string.Empty,
                                MaximumCoverageSequenceNo = workSheet.Cells[i, 2].Value != null ? workSheet.Cells[i, 2].Value.ToString() : string.Empty,
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
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
                    SalesmanCode = e.SalesManCode,
                    SalesmanName = e.SalesManCode,
                    SEType = e.SEType,
                    ReportingTo = e.ReportingToAM,
                    ReportingLevel = e.ReportingLevel,
                    SalesForceCode = e.SalesForceCode,
                    IsActive = e.IsActive,
                    Category = e.Category,
                    JoiningDate = e.JoiningDate,
                    Email = e.Email,
                    PhoneNo = e.MobileNumber1
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
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\Salesman.xlsx"));
            }
            catch (Exception ex)
            { }
            return this.folderPath + @"\Salesman.xlsx";
        }

        public void CreateCustomerRouteFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<CustomerRouteModel> rcpCustomerRoute = ReadRCPFromSalesTeam.Select(e => new CustomerRouteModel
                {
                    DistributorCode = e.DistributedBranchCode,
                    OutletCode = e.OutletID,
                    RouteCode = e.RouteCode,
                    MaximumCoverageSequenceNo = e.CoverageSequence
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
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\CustomerRoute.xlsx"));
            }
            catch (Exception ex)
            { }
        }

        public void CreateOutlet(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            try
            {
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sample sheet");
                List<OutletModel> outletModelList = new List<OutletModel>();
                bool isOlderExcel = IsOlderVersionExcel("Outlet");
                if (isOlderExcel)
                {
                    FillOutlet(outletModelList, ReadRCPFromSalesTeam);
                    CreateOutletFileHeader(ws);
                    for (int i = 0; i < outletModelList.Count; i++)
                    {
                        CreateOutletRawFileRow(ws, outletModelList, i);
                    }
                    ws.Protection.IsProtected = false;
                    ws.Protection.AllowSelectLockedCells = false;
                    ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\Outlet.xls"));
                }
                else
                    CreateOutletFile(ReadRCPFromSalesTeam);
            }
            catch (Exception ex)
            { }
        }

        public void CreateOutletFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<OutletModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new OutletModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    RetlrType = e.RetlrType,
                    OutletCode = e.OutletID,
                    OutletName = e.OutletName,
                    OutletAddress1 = e.OutletAddress,
                    Country = "India",
                    State = e.State,
                    City = e.City,
                    PostalCode = e.PostalCode,
                    Email = e.Email,
                    PhoneNo = e.MobileNumber1,
                    EnrollDate = string.Empty,
                    TaxType = "VAT",
                    CreditBills = string.Empty,
                    CreditBillAct = string.Empty,
                    CustChannelType = e.CustChannelType,
                    CustChannelSubType = e.CustChannelSubType,
                    IsActive = "Yes",
                    CreditDaysAct = "None",
                    CreditLimitAct = "None",
                    CashDiscPerc = "0",
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
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\Outlet.xlsx"));
            }
            catch (Exception ex)
            { }
        }

        public void CreateSalesmanRouteFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<SalesmanRouteModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new SalesmanRouteModel
                {
                    DistributorBranchCode = e.DistributedBranchCode,
                    SalesmanCode = e.SalesManCode,
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
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\SalesmanRoute.xlsx"));
            }
            catch (Exception ex)
            { }
        }

        public void CreateCustomerProductCategoryFile(List<RCPFromSalesTeamModel> ReadRCPFromSalesTeam)
        {
            try
            {
                List<CustomerProductCategoryModel> rcpOutlet = ReadRCPFromSalesTeam.Select(e => new CustomerProductCategoryModel
                {
                    OutletCode = e.OutletID,
                    ProductHierarchyLevelCode = "100",
                    ProductHierarchyValueCode = e.ProductHierarchyValueCode,
                    ProductHierarchyCategoryName = e.ProductHierarchyCategoryName
                }).ToList();

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet ws = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "OutletCode";
                ws.Cells[1, 2].Value = "ProductHierarchyLevelCode";
                ws.Cells[1, 3].Value = "ProductHierarchyValueCode";
                ws.Cells[1, 3].Value = "ProductHierarchyCategoryName";

                for (int i = 2; i < rcpOutlet.Count; i++)
                {
                    ws.Cells[i, 1].Value = rcpOutlet[i].OutletCode;
                    ws.Cells[i, 2].Value = rcpOutlet[i].ProductHierarchyLevelCode;
                    ws.Cells[i, 3].Value = rcpOutlet[i].ProductHierarchyValueCode;
                    ws.Cells[i, 4].Value = rcpOutlet[i].ProductHierarchyCategoryName;
                }
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\CustomerProductCategory.xlsx"));
            }
            catch (Exception ex)
            { }
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
                    Distance = "50",
                    Population = "50",
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
                ExcelPkg.SaveAs(new FileInfo(this.folderPath + @"\Beat.xlsx"));
            }
            catch (Exception ex)
            { }
        }
    }
}
