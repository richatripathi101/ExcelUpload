using InchesExcel.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.OleDb;
using InchesExcel.Models;
using Microsoft.EntityFrameworkCore;
using MySqlX.XDevAPI;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Wordprocessing;
using ClosedXML.Excel;
using System.Reflection;
using System.Globalization;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using Grpc.Core;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace InchesExcel.Controllers
{
    public class ClientExcelController : Controller

    {
        //string connectionString = ConnectionStrings.DefaultConnection;
        private readonly IConfiguration configuration;
        private readonly ApplicationContext context;

        private readonly string wwwrootDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");

        public IActionResult Template()
        {
            List<string> templ = Directory.GetFiles(wwwrootDirectory, "*.xlsx")
                .Select(Path.GetFileName).ToList();
            return View(templ);
        }
        [HttpPost]
        public async Task<IActionResult> Template(IFormFile myFile)
        {
            if (myFile != null)
            {
                var path = Path.Combine(
                    wwwrootDirectory,
                    DateTime.Now.Ticks.ToString() + Path.GetExtension(myFile.FileName));

                using (var stream = new FileStream(path, FileMode.Create))
                {
                    await myFile.CopyToAsync(stream);
                }
            }
            return View();
        }
        public async Task<IActionResult> DownloadFile(string filePath)
        {
            var path = Path.Combine(
                   Directory.GetCurrentDirectory(),
                   "wwwroot", filePath);
            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            var contentType = "APPLICATION/octet-stream";
            var fileName = Path.GetFileName(path);

            return File(memory, contentType, fileName);
        }
        public ClientExcelController(IConfiguration configuration, ApplicationContext context) 
        {
            this.configuration = configuration;
            this.context = context;
        }
       
       
       
        public IActionResult Index(int? pageNumber)
        {
            bager();
            int ? pageSize = 10;
            return View(PaginatedList<ClientExcel>.Create(context.ClientExcel.ToList(), pageNumber ?? 1, (int)pageSize));
            
        }

        

        //GET
        public IActionResult ImportExcelFile()
        {
            return View();
        }
        [HttpPost]
        public IActionResult ImportExcelFile(IFormFile formFile)
        { 
            try
            {
                
                
                    var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
                    if (!Directory.Exists(mainPath))
                    {
                        Directory.CreateDirectory(mainPath);
                    }
                    var filePath = Path.Combine(mainPath, formFile.FileName);
                    using (FileStream stream = new FileStream(filePath, FileMode.Create))
                    {
                        formFile.CopyTo(stream);
                    }
                    var fileName = formFile.FileName;
                    string extension = Path.GetExtension(fileName);
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls":
                            conString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                        case ".xlsx":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                    }
                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);
                    using (OleDbConnection conExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = conExcel;     
                                conExcel.Open();
                                DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["Table_Name"].ToString();
                                cmdExcel.CommandText = "SELECT * FROM    [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                conExcel.Close();

                            }
                        }
                    }
                    conString = configuration.GetConnectionString("DefaultConnection");
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            sqlBulkCopy.DestinationTableName = "ClientExcel";
                            sqlBulkCopy.ColumnMappings.Add("policy number", "Policynumber");
                            sqlBulkCopy.ColumnMappings.Add("UW Name", "UWName");
                            sqlBulkCopy.ColumnMappings.Add("Date", "Date");
                            sqlBulkCopy.ColumnMappings.Add("Lot", "Lot");
                            sqlBulkCopy.ColumnMappings.Add("Received timing", "Receivedtiming");
                            sqlBulkCopy.ColumnMappings.Add("Done Cases", "DoneCases");
                            sqlBulkCopy.ColumnMappings.Add("TAT", "TAT");
                            sqlBulkCopy.ColumnMappings.Add("With in TAT", "WithinTAT");
                            sqlBulkCopy.ColumnMappings.Add("Cases Status", "CasesStatus");
                            con.Open();
                            sqlBulkCopy.WriteToServer(dt);
                            con.Close();
                        }
                    }
                    TempData["message"] = "File Imported Successfully, Data Saved into Database ";
                    return RedirectToAction("Index");
                
                
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
            }
           
            return View();
        }
        public void bager()
        {
            ViewBag.displayClientExcel = context.ClientExcel.ToList();
            ViewBag.Count = context.ClientExcel.Count();
        }
        public IActionResult Delete(int id)
        {
            var emp = context.ClientExcel.SingleOrDefault(e => e.Id == id);
            context.ClientExcel.Remove(emp);
            context.SaveChanges();
            return RedirectToAction("Index");
        }

        //GET 
        public IActionResult Edit(int id)
        {
            var emp = context.ClientExcel.FirstOrDefault(e => e.Id == id);
            ClientExcel c = new ClientExcel();
            if (emp != null)
            {
                c.Id = emp.Id;
                c.Policynumber = emp.Policynumber;
                c.UWName = emp.UWName;
                c.Date = emp.Date;
                c.Lot = emp.Lot;
                c.Receivedtiming = emp.Receivedtiming;
                c.DoneCases = emp.DoneCases;
                c.TAT = emp.TAT;
                c.WithinTAT = emp.WithinTAT;
                c.CasesStatus = emp.CasesStatus;
            }

            return View(c);
        }
        [HttpPost]
        public IActionResult Edit(ClientExcel model)
        {
            var emp = new ClientExcel()
            {
                Id = model.Id,
                Policynumber = model.Policynumber,
                UWName = model.UWName,
                Date = model.Date,
                Lot = model.Lot,
                Receivedtiming = model.Receivedtiming,
                DoneCases = model.DoneCases,
                TAT = model.TAT,
                WithinTAT = model.WithinTAT,
                CasesStatus = model.CasesStatus

            };
            context.ClientExcel.Update(emp);
            context.SaveChanges();
            return RedirectToAction("Index");
        }

        public IActionResult ExportExcel()
        {
            try
            {
                var data = context.ClientExcel.ToList();
                if (data !=null & data.Count >0)
                {
                    using(XLWorkbook wb=new XLWorkbook())
                    {
                        wb.Worksheets.Add(ToConvertDataTable(data.ToList()));
                        using(MemoryStream stream=new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            string fileName = $"ClientExcel_{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocuments.spreadsheetml.sheet", fileName);
                        }
                    }
                }
                TempData["Error"] = "Data not found!";
            }
            catch(Exception ex)
            {

            }
            return RedirectToAction("Index");
        }


        public DataTable ToConvertDataTable<T>(List<T> items)
        {
            DataTable dt =new DataTable(typeof(T).Name);
            PropertyInfo[] propInfo= typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in propInfo) 
            {
                dt.Columns.Add(prop.Name); 
            }
            foreach (T item in items)
            {
                var values = new object[propInfo.Length];
                for (int i = 0; i < propInfo.Length; i++)
                {
                    values[i] = propInfo[i].GetValue(item, null);
                }
                dt.Rows.Add(values);
            }          
            return dt;
        }

        public async Task<IActionResult> Search(String SearchString)
        {
            ViewData["CurrentFilter"] = SearchString;
            var pol = from b in context.ClientExcel select b;
            if(!String.IsNullOrEmpty(SearchString) ) 
            {
                pol = pol.Where(b => b.Policynumber.Contains(SearchString));
            }
           return View(pol);
        }

        [HttpPost]
        public IActionResult SearchBetweenDates(DateTime start, DateTime end)
        {
            string conString = string.Empty;
            List<ClientExcel> client = new List<ClientExcel>();
            conString = configuration.GetConnectionString("DefaultConnection");
            using (SqlConnection con = new SqlConnection(conString))
            {
                SqlCommand cmd = new SqlCommand("betweenDates", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@start", start);
                cmd.Parameters.AddWithValue("@end", end);
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    ClientExcel model = new ClientExcel();
                    model.Id = (int)rdr["Id"];
                    model.Policynumber = rdr["Policynumber"].ToString();
                    model.UWName = rdr["UWName"].ToString();
                    model.Date = (DateTime)rdr["Date"];
                    model.Lot = rdr["Lot"].ToString();
                    model.Receivedtiming = (DateTime)rdr["Receivedtiming"];
                    model.DoneCases = (DateTime)rdr["DoneCases"];
                    model.TAT = (int)rdr["TAT"];
                    model.WithinTAT = rdr["WithinTAT"].ToString();
                    model.CasesStatus = rdr["CasesStatus"].ToString();
                    client.Add(model);
                }
                con.Close();
                
            }
            try
            {
                // var data = context.ClientExcel.ToList();
                var data = client;
                if (data != null & data.Count > 0)
                {
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(ToConvertDataTable(data.ToList()));
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            string fileName = $"ClientExcel_{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocuments.spreadsheetml.sheet", fileName);
                        }
                    }
                }
                TempData["Error"] = "Data not found!";
            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }
    }
}


