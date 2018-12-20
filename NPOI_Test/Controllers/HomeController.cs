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
            //再次使用web git 提交修改
            //再次使用vs提交修改

            //继续使用web git 提交代码02
            //继续使用vs提交修改代码03
            //继续使用web git 提交代码03

            //如果本地提交之后立即推送，可能出现的情况：
            // 1 当没有其他人修改同一文件时，接下来的拉取操作正常执行。
            // 2 当有其他人修改了同一文件时，会进入冲突管理流程。

            //看看冲突是否可以利用先拉取得方式处理。来自vs的提交
            
            //看看冲突是否可以利用先拉取得方式处理。来自vs的提交
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile()
        {
            //last again 05 使用vs提交
            //again web git04
            //again web git03
            //again web git02
            //again web git
            //继续使用vs提交修改代码02
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
        ///     直接导出
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
        /// datatable转换成excel文件，并保存到本地服务器
        /// </summary>
        public void SaveFileByLocal()
        {
            //1.创建EXCEL中的Workbook  
            HSSFWorkbook myHSSFworkbook = new HSSFWorkbook();
            //2.创建Workbook中的Sheet  
            NPOI.SS.UserModel.ISheet mysheetHSSF = myHSSFworkbook.CreateSheet("sheet1");
            //3.创建Sheet中的Row  
            NPOI.SS.UserModel.IRow rowHSSF = mysheetHSSF.CreateRow(0);
            // SetCellValue 有5个重载方法 bool、DateTime、double、string、IRichTextString(未演示)  
            rowHSSF.CreateCell(0).SetCellValue(true);
            rowHSSF.CreateCell(1).SetCellValue(System.DateTime.Now);
            rowHSSF.CreateCell(2).SetCellValue(9.32);
            rowHSSF.CreateCell(3).SetCellValue("Hello World！");

            //5.保存  
            using (FileStream fileHSSF = new FileStream(@"E:\myHSSFworkbook.xls", FileMode.Create))
            {
                myHSSFworkbook.Write(fileHSSF);
            }
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
