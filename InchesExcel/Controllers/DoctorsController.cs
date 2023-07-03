using ExcelDataReader;
using InchesExcel.Data;
using InchesExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MongoDB.Driver.Core.Configuration;
using System.Data;


namespace InchesExcel.Controllers
{
    public class DoctorsController : Controller
    {
        private readonly IConfiguration configuration;
        private readonly ApplicationContext context;
        private readonly string wwwrootDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");

        public DoctorsController(IConfiguration configuration, ApplicationContext context)
        {
            this.configuration = configuration;
            this.context = context;
        }

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
        public IActionResult Index(int? pageNumber)
        {
            // bager();
            int? pageSize = 10;
            return View(PaginatedList<Doctors>.Create(context.Doctors.ToList(), pageNumber ?? 1, (int)pageSize));
            //var data = context.Doctors.ToList();
            //return View(data);
        }
        //GET
        //public IActionResult ImportExcelFile()
        //{
        //    return View();
        //}

        public ActionResult ImportExcelFile()
        {
            //DataTable dt = new DataTable();

            //if ((String)Session["tmpdata"] != null)
            //{
            try
            {
                //dt = (DataTable)Session["tmpdata"];
            }
            catch (Exception ex)
            {

            }
            //}


            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ImportExcelFile(DoctorViewModel model)
        {

            if (ModelState.IsValid)
            {

                if (model.formFile != null && model.formFile.Length > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = model.formFile.OpenReadStream();

                    // We return the interface, so that
                    IExcelDataReader reader = null;


                    if (model.formFile.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (model.formFile.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }
                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;


                    DataTable dt = new DataTable();
                    //dt.Columns.Add("UserName");
                    //dt.Columns.Add("Adddress");
                    DataRow row;


                    DataTable dt_ = new DataTable();
                    try
                    {

                        dt_ = reader.AsDataSet().Tables[0];

                        string ret = "";



                        for (int i = 0; i < dt_.Columns.Count; i++)
                        {
                            dt.Columns.Add(dt_.Rows[0][i].ToString());
                        }

                        int rowcounter = 0;
                        for (int row_ = 1; row_ < dt_.Rows.Count; row_++)
                        {
                            row = dt.NewRow();

                            for (int col = 0; col < dt_.Columns.Count; col++)
                            {
                                row[col] = dt_.Rows[row_][col].ToString();
                                rowcounter++;
                            }
                            dt.Rows.Add(row);
                        }

                    }
                    catch (Exception ex)
                    {
                        ModelState.AddModelError("File", "Unable to Upload file!");
                        return View();
                    }

                    DataSet result = new DataSet();//reader.AsDataSet();
                    result.Tables.Add(dt);
                    string minutes_ID = "";



                    reader.Close();
                    reader.Dispose();
                    // return View();
                    //  return View(result.Tables[0]);

                    DataTable ddd = result.Tables[0];
                    string conString = string.Empty;
                    conString = configuration.GetConnectionString("DefaultConnection");
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            sqlBulkCopy.DestinationTableName = "TempNonTmsCases";
                            sqlBulkCopy.ColumnMappings.Add("Date", "Date");
                            sqlBulkCopy.ColumnMappings.Add("Policy number", "PolicyNumber");
                            sqlBulkCopy.ColumnMappings.Add("SAR", "SAR");
                            sqlBulkCopy.ColumnMappings.Add("decision match with initial UW", "DecisionMatchWithInitialUW");
                            sqlBulkCopy.ColumnMappings.Add("Decision", "Decision");
                            sqlBulkCopy.ColumnMappings.Add("Remarks", "Remarks");
                            sqlBulkCopy.ColumnMappings.Add("CP details & Findings", "CPDetailAndFindings");
                            sqlBulkCopy.ColumnMappings.Add("Inches UW NAME", "InchesUWName");
                            sqlBulkCopy.ColumnMappings.Add("final decision", "FinalDecision");
                            sqlBulkCopy.ColumnMappings.Add("LOT SENT TIME", "LOTSentTime");
                            sqlBulkCopy.ColumnMappings.Add("Received Time", "ReceivedTime");
                            sqlBulkCopy.ColumnMappings.Add("TAT Status", "TATStatus");
                            sqlBulkCopy.ColumnMappings.Add("Team - DUW / CMO / QC", "Team");
                            sqlBulkCopy.ColumnMappings.Add("System Status", "SystemStatus");
                            sqlBulkCopy.ColumnMappings.Add("Case Type", "CaseType");

                            con.Open();
                            sqlBulkCopy.WriteToServer(ddd);
                            con.Close();

                        }
                    }
                    //Session["tmpdata"] = ddd;

                    return RedirectToAction("ImportExcelFile");

                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
        }




        //public IActionResult ImportExcelFile(IFormFile formFile)
        //{
        //    try
        //    {


        //        var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
        //        if (!Directory.Exists(mainPath))
        //        {
        //            Directory.CreateDirectory(mainPath);
        //        }
        //        var filePath = Path.Combine(mainPath, formFile.FileName);
        //        using (FileStream stream = new FileStream(filePath, FileMode.Create))
        //        {
        //            formFile.CopyTo(stream);
        //        }
        //        var fileName = formFile.FileName;
        //        string extension = Path.GetExtension(fileName);
        //        string conString = string.Empty;
        //        switch (extension)
        //        {
        //            case ".xls":
        //                conString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
        //                break;
        //            case ".xlsx":
        //                conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
        //                break;
        //        }
        //        DataTable dt = new DataTable();
        //        conString = string.Format(conString, filePath);
        //        using (OleDbConnection conExcel = new OleDbConnection(conString))
        //        {
        //            using (OleDbCommand cmdExcel = new OleDbCommand())
        //            {
        //                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
        //                {
        //                    cmdExcel.Connection = conExcel;
        //                    conExcel.Open();
        //                    DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //                    string sheetName = dtExcelSchema.Rows[0]["Table_Name"].ToString();
        //                    cmdExcel.CommandText = "SELECT * FROM    [" + sheetName + "]";
        //                    odaExcel.SelectCommand = cmdExcel;
        //                    odaExcel.Fill(dt);
        //                    conExcel.Close();

        //                }
        //            }
        //        }
        //        conString = configuration.GetConnectionString("DefaultConnection");
        //        using (SqlConnection con = new SqlConnection(conString))
        //        {
        //            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
        //            {
        //                sqlBulkCopy.DestinationTableName = "Doctors";
        //                sqlBulkCopy.ColumnMappings.Add("Date", "Date");
        //                sqlBulkCopy.ColumnMappings.Add("Policy number", "PolicyNumber");
        //                sqlBulkCopy.ColumnMappings.Add("decision match with initial UW", "DecisionMatchWithInitialUW");
        //                sqlBulkCopy.ColumnMappings.Add("Decision", "Decision");
        //                sqlBulkCopy.ColumnMappings.Add("Remarks", "Remarks");
        //                sqlBulkCopy.ColumnMappings.Add("CP details & Findings", "CPDetailAndFindings");
        //                sqlBulkCopy.ColumnMappings.Add("Inches UW NAME", "InchesUWname");
        //                sqlBulkCopy.ColumnMappings.Add("final decision", "FinalDecision");
        //                sqlBulkCopy.ColumnMappings.Add("LOT SENT TIME", "LOTsenttime");
        //                sqlBulkCopy.ColumnMappings.Add("Received Time", "Receivedtime");
        //                sqlBulkCopy.ColumnMappings.Add("TAT Status", "TATstatus");
        //                sqlBulkCopy.ColumnMappings.Add("Team - DUW / CMO / QC", "Team");
        //                sqlBulkCopy.ColumnMappings.Add("System Status", "SystemStatus");
        //                sqlBulkCopy.ColumnMappings.Add("Case Type", "CaseType");

        //                con.Open();
        //                sqlBulkCopy.WriteToServer(dt);
        //                con.Close();
        //            }
        //        }
        //        TempData["message"] = "File Imported Successfully, Data Saved into Database.";
        //        return RedirectToAction("Index");


        //    }
        //    catch (Exception ex)
        //    {
        //        string msg = ex.Message;
        //    }

        //    return View();
        //}
        public void bager()
        {
            ViewBag.displayDoctors = context.Doctors.ToList();
            ViewBag.Count = context.Doctors.Count();
        }
        public IActionResult Delete(int id)
        {
            var emp = context.Doctors.SingleOrDefault(e => e.Id == id);
            context.Doctors.Remove(emp);
            context.SaveChanges();
            return RedirectToAction("Index");
        }

        //GET 
        public IActionResult Edit(int id)
        {
            var emp = context.Doctors.FirstOrDefault(e => e.Id == id);
            Doctors c = new Doctors();
            if (emp != null)
            {
                c.Id = emp.Id;
                c.Date = emp.Date;
                c.PolicyNumber = emp.PolicyNumber;
                c.DecisionMatchWithInitialUW = emp.DecisionMatchWithInitialUW;
                c.Decision = emp.Decision;
                c.Remarks = emp.Remarks;
                c.CPDetailAndFindings = emp.CPDetailAndFindings;
                c.InchesUWname = emp.InchesUWname;
                c.FinalDecision = emp.FinalDecision;
                c.LOTsenttime = emp.LOTsenttime;
                c.Receivedtime = emp.Receivedtime;
                c.TATstatus = emp.TATstatus;
                c.Team = emp.Team;
                c.SystemStatus = emp.SystemStatus;
                c.CaseType = emp.CaseType;
            }

            return View(c);
        }
        [HttpPost]
        public IActionResult Edit(Doctors model)
        {
            var emp = new Doctors()
            {
                Id = model.Id,
                Date = model.Date,
                PolicyNumber = model.PolicyNumber,
                DecisionMatchWithInitialUW = model.DecisionMatchWithInitialUW,
                Decision = model.Decision,
                Remarks = model.Remarks,
                CPDetailAndFindings = model.CPDetailAndFindings,
                InchesUWname = model.InchesUWname,
                FinalDecision = model.FinalDecision,
                LOTsenttime = model.LOTsenttime,
                Receivedtime = model.Receivedtime,
                TATstatus = model.TATstatus,
                Team = model.Team,
                SystemStatus = model.SystemStatus,
                CaseType = model.CaseType

            };
            context.Doctors.Update(emp);
            context.SaveChanges();
            return RedirectToAction("Index");
        }

        //public IActionResult ExportExcel()
        //{
        //    try
        //    {
        //        var data = context.Doctors.ToList();
        //        if (data != null & data.Count > 0)
        //        {
        //            using (XLWorkbook wb = new XLWorkbook())
        //            {
        //                wb.Worksheets.Add(ToConvertDataTable(data.ToList()));
        //                using (MemoryStream stream = new MemoryStream())
        //                {
        //                    wb.SaveAs(stream);
        //                    string fileName = $"Doctors_{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";
        //                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocuments.spreadsheetml.sheet", fileName);
        //                }
        //            }
        //        }
        //        TempData["Error"] = "Data not found!";
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return RedirectToAction("Index");
        //}


        //public DataTable ToConvertDataTable<T>(List<T> items)
        //{
        //    DataTable dt = new DataTable(typeof(T).Name);
        //    PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        //    foreach (PropertyInfo prop in propInfo)
        //    {
        //        dt.Columns.Add(prop.Name);
        //    }
        //    foreach (T item in items)
        //    {
        //        var values = new object[propInfo.Length];
        //        for (int i = 0; i < propInfo.Length; i++)
        //        {
        //            values[i] = propInfo[i].GetValue(item, null);
        //        }
        //        dt.Rows.Add(values);
        //    }
        //    return dt;
        //}
        public async Task<IActionResult> Search(String SearchString)
        {
            ViewData["CurrentFilter"] = SearchString;
            var pol = from b in context.Doctors select b;
            if (!String.IsNullOrEmpty(SearchString))
            {
                pol = pol.Where(b => b.PolicyNumber.Contains(SearchString));
            }
            return View(pol);
        }

       
        //public IActionResult SearchBetweenDates(DateTime start, DateTime end)

        //{
        //    string conString = string.Empty;
        //    List<Doctors> doctors = new List<Doctors>();
        //    conString = configuration.GetConnectionString("DefaultConnection");
        //    using (SqlConnection con = new SqlConnection(conString))
        //    {
        //        SqlCommand cmd = new SqlCommand("betweenDatesDoc", con);
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.Parameters.AddWithValue("@start", start);
        //        cmd.Parameters.AddWithValue("@end", end);
        //        con.Open();
        //        SqlDataReader rdr = cmd.ExecuteReader();

        //        while (rdr.Read())
        //        {
        //            Doctors model = new Doctors();
        //            model.Id = (int)rdr["Id"];
        //            model.Date = (DateTime)rdr["Date"];
        //            model.PolicyNumber = rdr["PolicyNumber"].ToString();
        //            model.DecisionMatchWithInitialUW = rdr["DecisionMatchWithInitialUW"].ToString();
        //            model.Decision = rdr["Decision"].ToString();
        //            model.Remarks = rdr["Remarks"].ToString();
        //            model.CPDetailAndFindings = rdr["CPDetailAndFindings"].ToString();
        //            model.InchesUWname = rdr["InchesUWname"].ToString();
        //            model.FinalDecision = rdr["FinalDecision"].ToString();
        //            model.LOTsenttime = (DateTime)rdr["LOTsenttime"];
        //            model.Receivedtime = (DateTime)rdr["Receivedtime"];
        //            model.TATstatus = rdr["TATstatus"].ToString();
        //            model.Team = rdr["Team"].ToString();
        //            model.SystemStatus = rdr["SystemStatus"].ToString();
        //            model.CaseType = rdr["CaseType"].ToString();
        //            doctors.Add(model);
        //        }
        //        con.Close();

        //    }
        //    try
        //    {

        //        var data = doctors;
        //        if (data != null & data.Count > 0)
        //        {
        //            using (XLWorkbook wb = new XLWorkbook())
        //            {
        //                wb.Worksheets.Add(ToConvertDataTable(data.ToList()));
        //                using (MemoryStream stream = new MemoryStream())
        //                {
        //                    wb.SaveAs(stream);
        //                    string fileName = $"Doctors{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";
        //                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocuments.spreadsheetml.sheet", fileName);
        //                }
        //            }
        //        }
        //        TempData["Error"] = "Data not found!";
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return RedirectToAction("Index");



        //}

    
}
}

