using NPOI.HSSF.UserModel;
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
        /// NPOI方式导出excel
        /// </summary>
        /// <returns></returns>
        public FileResult ExportStu2()
        {
            string schoolname = "401";

            //创建Excel文件的对象
            HSSFWorkbook book = new HSSFWorkbook();

            //添加一个sheet
            NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet("Sheet1");

            //假数据(真实数据需要从数据库中获取)
            List<Compture> listRainInfo = new List<Compture>()
            {
                new Compture() { PCName="pc1",UserName="小明1" },
                new Compture() { PCName="pc2",UserName="小明2" },
                new Compture() { PCName="pc3",UserName="小明3"},
                new Compture() { PCName="pc4",UserName="小明4"},
            };

            //给sheet1添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(0);
            row1.CreateCell(0).SetCellValue("电脑号");
            row1.CreateCell(1).SetCellValue("姓名");

            //将数据逐步写入sheet1各个行
            for (int i = 0; i < listRainInfo.Count; i++)
            {
                NPOI.SS.UserModel.IRow rowtemp = sheet1.CreateRow(i + 1);
                rowtemp.CreateCell(0).SetCellValue(listRainInfo[i].PCName.ToString());
                rowtemp.CreateCell(1).SetCellValue(listRainInfo[i].UserName.ToString());
            }
            // 写入到客户端 
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return File(ms, "application/vnd.ms-excel", "第一批电脑派位生名册.xls");
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

    public class Compture
    {
        public string PCName { get; set; }
        public string UserName { get; set; }
    }
}