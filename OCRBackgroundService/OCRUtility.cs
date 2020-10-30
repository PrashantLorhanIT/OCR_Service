using Dynamsoft.Core;
using Dynamsoft.OCR;
using Dynamsoft.OCR.Enums;
using Dynamsoft.PDF;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Dynamsoft.WPF;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using Dynamsoft.Core.Enums;
using System.Net.Http;
using Newtonsoft.Json;
using log4net;
using System.Linq.Expressions;
using TikaOnDotNet.TextExtraction;
//using de.softpro.doc;

namespace OCRBackgroundService
{
    public class OCRUtility : IConvertCallback
    {
        private string m_strCurrentDirectory;
        private string m_StrProductKey = ConfigurationManager.AppSettings["ProductKey"];      
        private ImageCore m_ImageCore = null;
        private Tesseract m_Tesseract = null;
        private PDFRasterizer m_PDFRasterizer = null;
        Dictionary<string, string> languages = new Dictionary<string, string>();
        string[] resultFormat = new string[] { "Text File", "Adobe PDF Plain Text File", "Adobe PDF Image Over Text File" };
        List<Bitmap> tempListSelectedBitmap = null;
        OcrResultModel ocrModel;
        string[] fileFormat;
        string[] ExcludeForOCR;       
       
        public OCRUtility()
        {
            languages.Add("English", "eng");
            m_strCurrentDirectory = Directory.GetCurrentDirectory();
        }
        
        public void DirectorySearch()
        {
            try
            {
                //SignDocDocumentLoader loader = new SignDocDocumentLoader();
                fileFormat = ConfigurationManager.AppSettings["FileFormats"].Split(';');
                ExcludeForOCR = ConfigurationManager.AppSettings["ExcludeForOCR"].Split(';');
                GetOcrFileForProcessing(5);                
            }
            catch (System.Exception ex)
            {
                throw ex;                           
            }
        }

        public void GetOcrFileForProcessing(int count)
        {
            LogUtility.WriteToFile("GetOcrFileForProcessin" + DateTime.Now);
            try
            {
                using (var client = new HttpClient())
                {
                    string ApiKey = ConfigurationManager.AppSettings["ApiKey"];
                    string baseUrl = ConfigurationManager.AppSettings["BaseUrl"];
                    var filterCount = ConfigurationManager.AppSettings["Filter"];
                    client.BaseAddress = new Uri(baseUrl);
                    var url = client.BaseAddress + string.Format("Attachment/GetPendingOCRAttachments?Count={0}&Filter={1}", 5,filterCount);
                    //var url = ConfigurationManager.AppSettings["GetOCR"] + count;
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.TryAddWithoutValidation("ApiKey", ApiKey);
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");
                   
                    HttpResponseMessage response = client.GetAsync(url).Result;
                    response.EnsureSuccessStatusCode();
                    var resp = response.Content.ReadAsStringAsync().Result;
                    LogUtility.WriteToFile("Response:" + resp);
                    ocrModel = JsonConvert.DeserializeObject<OcrResultModel>(resp);
                    //if (ocrModel.data.Count == 0)
                    //{
                    //    OcrAttchmentModel a = new OcrAttchmentModel();
                    //    a.ridAttachment = "462";
                    //    ocrModel.data.Add(a);
                    //}
                    if (ocrModel != null && ocrModel.data.Count > 0)
                    {
                        foreach (OcrAttchmentModel ocrAtt in ocrModel.data)
                        {
                            //E:\ERCMSAPI\UploadedFiles\Attachments\ridCorr_387\E0001-EGS-JBS-CR-XXXXX.pdf
                            // ocrAtt.documentumid = "D:\\Attachments\\ridCorr_444\\New Microsoft Excel Worksheet.xlsx";
                            //ridCorr_444\X0231-JBS-EGS-CL-XXXXX.pdf
                            //"D:\\Attachments\\RIDCORR_442\\X0231-JBS-EGS-CL-XXXXX.pdf";
                            //ocrAtt.documentumid = "D:\\Attachments\\RIDMOM_462\\X0231-S2A-GDM-RP-00001.pdf";
                            LogUtility.WriteToFile("Process File Name:" + ocrAtt.documentumid);
                                //"Aptiva minutes of meeting-18 May 2020.docx";
                            string ext = Path.GetExtension(ocrAtt.documentumid);
                            if (fileFormat.Contains(ext))
                            {
                                //string s = ocrAtt.documentumid;
                                //ocrAtt.documentumid = s.Replace(@"E:\ERCMSAPI\UploadedFiles\Attachments", @"D:\Attachments");
                                if (File.Exists(ocrAtt.documentumid))
                                {
                                    if (ExcludeForOCR.Contains(ext))
                                    {
                                        LogUtility.WriteToFile("Method Name:GenerateByteArray");
                                        GenerateByteArray(ocrAtt, ocrAtt.documentumid, ocrAtt.attachedfilename);
                                    }
                                    else
                                    {
                                        LogUtility.WriteToFile("Method Name:GenerateOCR");
                                        GenerateOCR(ocrAtt, ocrAtt.documentumid, ocrAtt.attachedfilename);
                                    }
                                }
                                else
                                {
                                    LogUtility.WriteToFile("Id:" + ocrAtt.ridAttachment.ToString());
                                    SaveOCRWithoutContent(ocrAtt);
                                }
                            }
                            else
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void GenerateOCR(OcrAttchmentModel ocrAttModel, string sourcePath, string fileName)  //, string targetPath
        {          
           
            byte[] sbytes = null;
            string languageFolder = ConfigurationManager.AppSettings["Tessdata"];
            LogUtility.WriteToFile("OCR Utility Process Path:" + languageFolder);
            //m_strCurrentDirectory + ConfigurationManager.AppSettings["Tessdata"];     // "\\OCR\\tessdata";
            m_Tesseract = new Tesseract(m_StrProductKey);
            m_ImageCore = new ImageCore();
            m_PDFRasterizer = new PDFRasterizer(m_StrProductKey);
            tempListSelectedBitmap = new List<System.Drawing.Bitmap>();
            m_Tesseract.TessDataPath = languageFolder;
            m_Tesseract.Language = languages["English"];
            m_Tesseract.ResultFormat = ResultFormat.Text;          

            string imageFolder = sourcePath;          
            string testPath = string.Empty;

            if (sourcePath.Contains(".pdf"))
            {
                m_PDFRasterizer.ConvertMode = Dynamsoft.PDF.Enums.EnumConvertMode.enumCM_AUTO;
                m_PDFRasterizer.ConvertToImage(imageFolder, "", 200, this as IConvertCallback);              
            }
            else
            {
                m_ImageCore.IO.LoadImage(imageFolder);
            }
            int imgCount = m_ImageCore.ImageBuffer.HowManyImagesInBuffer;
            List<short> lstindexs = new List<short>();
            for (short index = 0; index < imgCount; index++)
            {
                if (index >= 0 && index < m_ImageCore.ImageBuffer.HowManyImagesInBuffer)
                {
                    if (tempListSelectedBitmap == null)
                    {
                        tempListSelectedBitmap = new List<Bitmap>();
                    }
                    Bitmap temp = m_ImageCore.ImageBuffer.GetBitmap(index);
                    tempListSelectedBitmap.Add(temp);
                    lstindexs.Add(index);
                }
            }

            if (tempListSelectedBitmap != null)
                sbytes = m_Tesseract.Recognize(tempListSelectedBitmap);

            if (sbytes != null && sbytes.Length > 0)
            {
                LogUtility.WriteToFile("OCR Processing File Name:" + fileName);
                SaveOCRBlob(ocrAttModel, fileName, sbytes);
            }
        }

        public void GenerateByteArray(OcrAttchmentModel ocrAttModel, string sourcePath, string fileName)
        {
            
            byte[] sbytes = null;
            sbytes = File.ReadAllBytes(ocrAttModel.documentumid);           
            if (sbytes != null && sbytes.Length > 0)  
            {                
                SaveOCRBlob(ocrAttModel, fileName, sbytes, true);
            }         
        }

        public void LoadConvertResult(ConvertResult result)
        {
            m_ImageCore.IO.LoadImage(result.Image);
            m_ImageCore.ImageBuffer.SetMetaData(m_ImageCore.ImageBuffer.CurrentImageIndexInBuffer, EnumMetaDataType.enumAnnotation, result.Annotations, true);
        }

        public void SaveOCRBlob(OcrAttchmentModel ocrAttModel, string fileName, Byte[] sbytes, bool isNotPdf = false )
        {
            using (var client = new HttpClient())
            {
                string ApiKey = ConfigurationManager.AppSettings["ApiKey"];
                string baseUrl = ConfigurationManager.AppSettings["BaseUrl"];
                client.BaseAddress = new Uri(baseUrl);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");
                client.DefaultRequestHeaders.TryAddWithoutValidation("ApiKey", ApiKey);
                LogUtility.WriteToFile("OCR Process Request PayLoad Header:" + client.DefaultRequestHeaders);
                //HTTP POST
                OcrBlobModel ocrBlobModel = new OcrBlobModel();
                ocrBlobModel.ridAttachment = Convert.ToInt32(ocrAttModel.ridAttachment);
                ocrBlobModel.contents = isNotPdf == true ? new TextExtractor().Extract(sbytes).Text : Encoding.UTF8.GetString(sbytes, 0, sbytes.Length);
                string strPayload = JsonConvert.SerializeObject(ocrBlobModel);
                LogUtility.WriteToFile("OCR Process Request PayLoad:" + strPayload);
                HttpContent c = new StringContent(strPayload, Encoding.UTF8, "application/json");
                string path1 = client.BaseAddress + "Attachment/CreateOCREntry";
                    //ConfigurationManager.AppSettings["CreateOCR"];                
                HttpResponseMessage response = client.PostAsync(path1, c).Result;
                LogUtility.WriteToFile("OCR Process Response:" + response);
            }
        }

        public void SaveOCRWithoutContent(OcrAttchmentModel ocrAttModel)
        {
            using (var client = new HttpClient())
            {
                string ApiKey = ConfigurationManager.AppSettings["ApiKey"];
                string baseUrl = ConfigurationManager.AppSettings["BaseUrl"];
                client.BaseAddress = new Uri(baseUrl);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");
                client.DefaultRequestHeaders.TryAddWithoutValidation("ApiKey", ApiKey);
                //HTTP POST
                OcrBlobModel ocrBlobModel = new OcrBlobModel();
                ocrBlobModel.ridAttachment = Convert.ToInt32(ocrAttModel.ridAttachment);
                ocrBlobModel.contents =" ";
                string strPayload = JsonConvert.SerializeObject(ocrBlobModel);
                HttpContent c = new StringContent(strPayload, Encoding.UTF8, "application/json");
                string path1 = client.BaseAddress + "Attachment/CreateOCREntry";
                    //ConfigurationManager.AppSettings["CreateOCR"];
                HttpResponseMessage response = client.PostAsync(path1, c).Result;
            }
        }
    }
}
