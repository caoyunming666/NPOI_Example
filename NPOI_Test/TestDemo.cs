using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace NPOI_Test
{
    public class TestDemo
    {
        static DataTable GenerateData()
        {
            DataTable data = new DataTable();
            for (int i = 0; i < 5; ++i)
            {
                data.Columns.Add("Columns_" + i.ToString(), typeof(string));
            }

            for (int i = 0; i < 10; ++i)
            {
                DataRow row = data.NewRow();
                row["Columns_0"] = "item0_" + i.ToString();
                row["Columns_1"] = "item1_" + i.ToString();
                row["Columns_2"] = "item2_" + i.ToString();
                row["Columns_3"] = "item3_" + i.ToString();
                row["Columns_4"] = "item4_" + i.ToString();
                data.Rows.Add(row);
            }
            return data;
        }

        static void PrintData(DataTable data)
        {
            if (data == null) return;
            for (int i = 0; i < data.Rows.Count; ++i)
            {
                for (int j = 0; j < data.Columns.Count; ++j)
                    Console.Write("{0} ", data.Rows[i][j]);
                Console.Write("\n");
            }
        }

        static void TestExcelWrite(string file)
        {
            try
            {
                using (ExcelHelper excelHelper = new ExcelHelper(file))
                {
                    DataTable data = GenerateData();
                    int count = excelHelper.DataTableToExcel(data, "MySheet", true);
                    if (count > 0)
                        Console.WriteLine("Number of imported data is {0} ", count);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }

        static void TestExcelRead(string file)
        {
            try
            {
                using (ExcelHelper excelHelper = new ExcelHelper(file))
                {
                    DataTable dt = excelHelper.ExcelToDataTable("MySheet", true);
                    PrintData(dt);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }

        static void Main(string[] args)
        {
            string file = "..\\..\\myTest.xlsx";
            TestExcelWrite(file);
            TestExcelRead(file);
        }
    }
}
