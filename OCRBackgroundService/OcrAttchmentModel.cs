using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;

namespace OCRBackgroundService
{
    public class OcrResultModel
    {
        public List<OcrAttchmentModel> data;
    }
    public class OcrAttchmentModel
    {
        public string ridAttachment { get; set; }
        public int attachsequence { get; set; }
        public string attachedfilename { get; set; }
        public string isactive { get; set; }
        public string addedby { get; set; }      
        public string addedon { get; set; }
        public string updatedby { get; set; }
        public string updatedon { get; set; }
        public string documentumid { get; set; }
        public int? ridCorr { get; set; }
        public int? ridRfi { get; set; }
        public int? ridMom { get; set; }
        public int? ridTask { get; set; }
     
    }

    public class OcrBlobModel
    {
        //public decimal? ridAttachmentocr { get; set; }
        public int ridAttachment { get; set; }
        // public string isactive { get; set; }
        public string contents { get; set; }
        // public string addedon { get; set; }

    }
}
