using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NPOI_Test.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile()
        {
            int success = 0;
            string message = "成功";

            string appPath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = appPath + "/Pub_Data";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            if (Request.Files.Count > 0)
            {
                HttpPostedFileBase fb = Request.Files["uploadObj"];
                string exName = Path.GetExtension(fb.FileName);
                string fileName = DateTime.Now.ToString("yyyyMMddHHmmss");
                string fileFullName = fileName + exName;
                string imgPath = filePath + "\\" + fileFullName;
                fb.SaveAs(imgPath);

                //读取excel文件数据
                TestExcelRead(imgPath);
            }
            return Json(new { success = success, message = message }, JsonRequestBehavior.DenyGet);
        }


        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="file"></param>
        public void TestExcelRead(string file)
        {
            try
            {
                using (ExcelHelper excelHelper = new ExcelHelper(file))
                {
                    DataTable dt = excelHelper.ExcelToDataTable("", true);

                    if (dt != null)
                    {
                        // do something...
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
    }
}