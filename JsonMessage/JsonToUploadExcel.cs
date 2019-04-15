using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.App_Code.QueueMessage
{
    public class JsonToUploadExcel : QDetails
    {
        #region Public Variables

        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string RegionCode { get; set; }
        public string TownCode { get; set; }
        public string FileType { get; set; }
        public string FileName { get; set; }
        public string Url { get; set; }

        #endregion

        public JsonToUploadExcel(string startDate, string endDate, string regionCode, string townCode, string fileType, string fileName, string url)
        {
            StartDate = startDate;
            EndDate = endDate;
            RegionCode = regionCode;
            TownCode = townCode;
            FileType = fileType;
            FileName = fileName;
            Url = url;
        }

        protected override void CreateFromJSON(string json)
        {
            //JObject pgmJObj = JObject.Parse(json);
            //JObject jObjectDetails = (JObject)pgmJObj[JSON.Tags.Message.Details.Key];
            //FromDate = jObjectDetails[JSON.Tags.Invoice.ProgramName].ToString();
            //FromDate = jObjectDetails[JSON.Tags.Invoice.ProgramName].ToString();
            //FromDate = jObjectDetails[JSON.Tags.Invoice.ProgramName].ToString();
        }

        protected override string DetailString()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>() { { "StartDate", StartDate }, { "EndDate", EndDate }, { "RegionCode", RegionCode }, { "TownCode", TownCode }, { "FileType", FileType }, { "FileName", FileName }, { "Url", Url } };
            var json = JsonConvert.SerializeObject(dict);
            return json;
        }

        protected override string DetailStringForBeat()
        {
            return "";
        }
    }
}
