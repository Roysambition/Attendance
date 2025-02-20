using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using AttendanceSystem.Models;

namespace AttendanceSystem.Controllers
{
    public class AttendanceController : Controller
    {
        private readonly string _connectionString;

        public AttendanceController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        // 1️⃣ 取得所有考勤紀錄
        public IActionResult Index()
        {
            List<Attendance> records = new List<Attendance>();

            using (MySqlConnection conn = new MySqlConnection(_connectionString))
            {
                conn.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT * FROM Attendance", conn);
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    records.Add(new Attendance
                    {
                        Id = reader.GetInt32("Id"),
                        EmployeeId = reader.GetString("EmployeeId"),
                        Name = reader.GetString("Name"),
                        Date = reader.GetDateTime("Date"),
                        CheckIn = reader.GetTimeSpan("CheckIn"),
                        CheckOut = reader.GetTimeSpan("CheckOut"),
                        WorkHours = reader.GetInt32("WorkHours")
                    });
                }
            }

            return View(records);
        }

        // 2️⃣ 匯出 Excel
        public IActionResult ExportToExcel()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("考勤報表");

                // 設定標題列
                IRow headerRow = sheet.CreateRow(0);
                string[] headers = { "員工編號", "姓名", "日期", "上班時間", "下班時間", "工時" };
                for (int i = 0; i < headers.Length; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(headers[i]);
                }

                // 取得考勤資料
                List<Attendance> records = new List<Attendance>();
                using (MySqlConnection conn = new MySqlConnection(_connectionString))
                {
                    conn.Open();
                    MySqlCommand cmd = new MySqlCommand("SELECT * FROM Attendance", conn);
                    MySqlDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        records.Add(new Attendance
                        {
                            EmployeeId = reader.GetString("EmployeeId"),
                            Name = reader.GetString("Name"),
                            Date = reader.GetDateTime("Date"),
                            CheckIn = reader.GetTimeSpan("CheckIn"),
                            CheckOut = reader.GetTimeSpan("CheckOut"),
                            WorkHours = reader.GetInt32("WorkHours")
                        });
                    }
                }

                // 寫入資料
                for (int i = 0; i < records.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(records[i].EmployeeId);
                    row.CreateCell(1).SetCellValue(records[i].Name);
                    row.CreateCell(2).SetCellValue(records[i].Date.ToString("yyyy-MM-dd"));
                    row.CreateCell(3).SetCellValue(records[i].CheckIn.ToString());
                    row.CreateCell(4).SetCellValue(records[i].CheckOut.ToString());
                    row.CreateCell(5).SetCellValue(records[i].WorkHours);
                }

                // 下載 Excel
                workbook.Write(ms);
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "考勤報表.xlsx");
            }
        }
    }
}
