using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormatting.Model
{
    public class RCPFromSalesTeamModel
    {
        public string RegionName { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string TownName { get; set; }
        public string ReportingTo { get; set; }
        public int SalesForceCode { get; set; }
        public string DistributedBranchCode { get; set; }
        public string RouteName { get; set; }
        public string IsNewRoute { get; set; }
        public string RouteCode { get; set; }
        public string SEType { get; set; }
        public string SECode { get; set; }
        public string IsNewSalesman { get; set; }        
        public string SalesmanCategory { get; set; }        
        public string DayOfWeek { get; set; }        
        public string OutletCode { get; set; }  
        public string OutletName { get; set; }
        public string OutletAddress { get; set; }        
        public string ProductHierarchyCategoryName { get; set; }
        public int PostalCode { get; set; }
        public string Retlrtype { get; set; }
        public string CustChannelType { get; set; }
        public string CustChannelSubType { get; set; }
        public long OutletPhoneNo { get; set; }
        public string StoreType { get; set; }
        public string OutletIdForRemoval { get; set; }
        public string RouteCodeforOutletIdRemoval { get; set; }
        public string OutletIdforRouteTransfer { get; set; }
        public string RouteCodeforOutletRouteTransfer { get; set; }
    }
}
