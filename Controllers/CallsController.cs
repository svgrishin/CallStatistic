using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using CallStatistic.Data;
using CallStatistic.Models;
using System.IO;
using ExcelDataReader;
using System.Data;
using Microsoft.SqlServer.Server;
using NuGet.Packaging;
using System.Collections.ObjectModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using Microsoft.VisualBasic;
//using Excel = Spire.Xls;
using System.Data.OleDb;
using System.Diagnostics;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using NuGet.Protocol.Plugins;
using CsvHelper;
using System.Globalization;
using Org.BouncyCastle.X509;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using rdr = SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;
using System.Text;
using Microsoft.EntityFrameworkCore.SqlServer.Query.Internal;
using NuGet.Protocol;
using Spire.Xls;
using Azure;
using SpreadsheetGear.Charts;
//using Spire.Xls;
//using OfficeOpenXml.Core.ExcelPackage;

namespace CallStatistic.Controllers
{
    public class CallsController : Controller
    {
        private readonly CallsContext _context;
        private DataTableCollection tableCollecton;

        public CallsController(CallsContext context)
        {
            _context = context;
        }

        // GET: Calls
        public async Task<IActionResult> Index()
        {
            return _context.Calls != null ?
                        View(await _context.Calls.ToListAsync()) :
                        Problem("Entity set 'CallsContext.Calls'  is null.");
        }

        // GET: Calls/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null || _context.Calls == null)
            {
                return NotFound();
            }

            var calls = await _context.Calls
                .FirstOrDefaultAsync(m => m.id == id);
            if (calls == null)
            {
                return NotFound();
            }

            return View(calls);
        }

        // GET: Calls/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Calls/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("id,dateOf,fromPhone,toPhone,duration")] Calls calls)
        {
            if (ModelState.IsValid)
            {
                _context.Add(calls);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(calls);
        }

        // GET: Calls/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null || _context.Calls == null)
            {
                return NotFound();
            }

            var calls = await _context.Calls.FindAsync(id);
            if (calls == null)
            {
                return NotFound();
            }
            return View(calls);
        }

        // POST: Calls/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("id,dateOf,fromPhone,toPhone,duration")] Calls calls)
        {
            if (id != calls.id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(calls);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CallsExists(calls.id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(calls);
        }

        // GET: Calls/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null || _context.Calls == null)
            {
                return NotFound();
            }

            var calls = await _context.Calls
                .FirstOrDefaultAsync(m => m.id == id);
            if (calls == null)
            {
                return NotFound();
            }

            return View(calls);
        }

        // POST: Calls/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            if (_context.Calls == null)
            {
                return Problem("Entity set 'CallsContext.Calls'  is null.");
            }
            var calls = await _context.Calls.FindAsync(id);
            if (calls != null)
            {
                _context.Calls.Remove(calls);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool CallsExists(int id)
        {
            return (_context.Calls?.Any(e => e.id == id)).GetValueOrDefault();
        }

        public IActionResult LoadCalls()
        {
            //SaveDataFromFiles("C:\\Users\\Сергей\\Desktop\\excelFiles\\");
            SaveDataFromFiles("excelFiles");
            //SaveDataFromFiles("C:\\excelFiles\\");

            _context.SaveChanges();

            return RedirectToAction("Index");
        }

        public void SaveDataFromFiles(string location)
        {



            string[] directories = Directory.GetDirectories(location);

            using (Process process = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                foreach (string directory in directories)
                {
                    string[] filesList = Directory.GetFiles(string.Concat(directory));
                    foreach (string file in filesList)
                    {
                        //string newFile = file;

                        string fileExtention = Path.GetExtension(file).ToUpper();

                        string newFile = file.Substring(0, file.IndexOf('.'));
                        newFile += ".xlsx";

                        FileStream stream;

                        info.FileName = "CMD.exe";
                        //info.WorkingDirectory = "C:\\Program Files (x86)\\Microsoft Office\\Office16";
                        info.WorkingDirectory = "C:\\Program Files\\Microsoft Office\\root\\Office16";
                        info.RedirectStandardInput = true;
                        info.UseShellExecute = false;
                        info.CreateNoWindow = true;
                        info.Arguments = "/C excelcnv.exe -oice \"" + file + "\" \"" + newFile + "\"";
                        process.StartInfo = info;

                        process.Start();
                        process.WaitForExit();
                        process.Close();



                        //string connString = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" + file + ";Extended Properties=Excel 5.0";
                        string connString = "Provider=Microsoft.ace.OLEDB.12.0;Data Source=" + newFile + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=2\"";

                        // Create the connection object 
                        OleDbConnection oledbConn = new OleDbConnection(connString);
                        //try
                        //{
                        // Open connection
                        oledbConn.Open();

                        System.Data.DataTable dataTable = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                        DataRow schemaRow = dataTable.Rows[0];
                        //string test = table.Rows[0].ItemArray[2].ToString(); 

                        string sheet = schemaRow["TABLE_NAME"].ToString();

                        OleDbCommand cmd = oledbConn.CreateCommand();
                        cmd.CommandText = "SELECT * FROM [" + sheet + "]";
                        //cmd.CommandText = "SELECT * FROM [" + test + "]";

                        OleDbDataReader reader = cmd.ExecuteReader();
                        System.Data.DataTable table = new System.Data.DataTable();
                        table.Load(reader);
                        //table.Load(reader);
                        //string s = dataTable.Rows[0][0].ToString();

                        string filesForDelete = null;

                        {



                            ///OldVersion
                            {
                                ////// Create OleDbCommand object and select data from worksheet Sheet1
                                ////OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);

                                //// Create new OleDbDataAdapter 
                                //OleDbDataAdapter oleda = new OleDbDataAdapter();

                                //oleda.SelectCommand = cmd;

                                //// Create a DataSet which will hold the data extracted from the worksheet.
                                //DataSet ds = new DataSet();

                                //// Fill the DataSet from the data extracted from the worksheet.
                                //oleda.Fill(ds);

                                ////// Bind the data to the GridView
                                ////GridView1.DataSource = ds.Tables[0].DefaultView;
                                ////GridView1.DataBind();
                            }

                            //}
                            //catch
                            //{
                            //}
                            //finally
                            //{
                            //    // Close connection
                            //    oledbConn.Close();
                            //}


                            ///TestVersion
                            {
                                //var workbook = rdr.Factory.GetWorkbook(file);
                                //var worksheet = workbook.Worksheets[0];
                                //var cells = worksheet.UsedRange;
                            }

                            ///WorkedVersion
                            {
                                //Excel.Workbook wb = x1.Workbooks.Open(file);
                                ////Workbook wb = x1.Workbooks.Open(file);
                                ////wb.LoadFromFile(file);
                            }

                            ///WorkedVersion
                            //newFile = file.Substring(0, file.IndexOf('.'));

                            ///WorkedVersion
                            {
                                //switch (fileExtention)
                                //{
                                //    case ".XLS":

                                //        ///TestVersion
                                //        {
                                //            //newFile = file.Substring(0, file.IndexOf('.'));
                                //            //newFile = newFile + "_new.txt";

                                //            //workbook.SaveAs(newFile, rdr.FileFormat.UnicodeText);
                                //            //workbook.Close();
                                //        }

                                //        ///WorkedVersion
                                //        {
                                //            //    //wb.SaveToFile(newFile);
                                //            //    wb.SaveAs(newFile, 51);
                                //            //    wb.Close();
                                //            //    x1.Quit();
                                //        }

                                //        process.StartInfo.Arguments = "/c del \"" + file + "\" /s";
                                //        process.Start();
                                //        process.Close();

                                //        break;

                                //    case ".CSV":

                                //        ///TestVersion
                                //        {
                                //            //workbook.SaveAs(newFile, rdr.FileFormat.OpenXMLWorkbook);
                                //            //workbook.Close();
                                //        }

                                //        ///WorkedVersion
                                //        {
                                //            //    wb.SaveAs(newFile, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                //            //    //wb.SaveToFile(newFile);


                                //            //wb.Close();
                                //            //x1.Quit();
                                //        }

                                //        process.StartInfo.Arguments = "/c del \"" + file + "\" /s";
                                //        process.Start();
                                //        process.Close();

                                //        break;

                                //    default:

                                //        ///WorkedVersion
                                //        {
                                //            //wb.Close();
                                //            //x1.Quit();
                                //        }

                                //        break;
                                //}

                                //stream = System.IO.File.Open(newFile, FileMode.Open, FileAccess.Read);

                                //IExcelDataReader excelReader;
                                //excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                                //var conf = new ExcelDataSetConfiguration
                                //{
                                //    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                //    {
                                //        UseHeaderRow = false
                                //    }
                                //};

                                //DataSet result = excelReader.AsDataSet(conf);

                                //System.Data.DataTable dt = result.Tables[0];
                            }
                        }

                        string folderName = directory.Substring(location.Length, directory.Length - location.Length);

                        switch (folderName)
                        {
                            case "ТТК":
                                //if (dt.Rows[0][0].ToString() == "№ договора")
                                if (table.Rows[0][0].ToString() == "№ договора")
                                {
                                    //dt.Rows[0].Delete();
                                    table.Rows[0].Delete();
                                    //foreach (DataRow row in dt.Rows)
                                    foreach (DataRow row in table.Rows)
                                    {
                                        if (row.RowState != DataRowState.Deleted)
                                        {
                                            string s = row[4].ToString();
                                            DateTime dateOf = DateTime.Parse(s);
                                            string fromPhone = row[2].ToString();
                                            string toPhone = row[7].ToString();

                                            string duration = row[8].ToString();
                                            if (duration.Length < 2) duration = "0" + duration;
                                            duration += ":00";

                                            Calls calls = new Calls(dateOf, fromPhone, toPhone, duration);
                                            _context.Calls.Add(calls);
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i <= 3; i++)
                                    {
                                        //dt.Rows[i].Delete();
                                        table.Rows[i].Delete();
                                    }
                                    //int rowCount = dt.Rows.Count;
                                    int rowCount = table.Rows.Count;
                                    for (int i = rowCount - 2; i < rowCount; i++)
                                    {
                                        //dt.Rows[i].Delete();
                                        table.Rows[i].Delete();
                                    }
                                    //foreach (DataRow row in dt.Rows)
                                    foreach (DataRow row in table.Rows)
                                    {
                                        if (row.RowState != DataRowState.Deleted)
                                        {
                                            string s = row[0].ToString();
                                            DateTime dateOf = DateTime.Parse(s);
                                            string fromPhone = row[1].ToString();
                                            string toPhone = row[2].ToString();

                                            int totalSeconds = Convert.ToInt16(row[3]);
                                            int minute = Convert.ToInt16(TimeSpan.FromSeconds(totalSeconds).TotalMinutes);
                                            int seconds = totalSeconds - minute * 60;

                                            string duration = minute.ToString();
                                            if (duration.Length < 2) duration = "0" + duration;
                                            string durationSeconds = seconds.ToString().Replace("-", "");
                                            if (durationSeconds.Length < 2) durationSeconds = "0" + durationSeconds;
                                            duration += ":" + durationSeconds;

                                            Calls calls = new Calls(dateOf, fromPhone, toPhone, duration);
                                            _context.Calls.Add(calls);
                                        }
                                    }
                                }
                                //excelReader.Close();

                                //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                //process.Start();
                                //process.Close();

                                reader.Close();
                                oledbConn.Close();

                                filesForDelete = file.Substring(0, file.IndexOf('.'));

                                info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                process.Start();
                                process.WaitForExit();
                                process.Close();

                                break;
                            case "Вымпелком":
                                //dt.Rows[0].Delete();
                                table.Rows[0].Delete();
                                //foreach (DataRow row in dt.Rows)
                                foreach (DataRow row in table.Rows)
                                {
                                    if (row.RowState != DataRowState.Deleted)
                                    {
                                        string s = row[0].ToString().Substring(0, 10);

                                        string s2 = TimeOnly.Parse(row[1].ToString()).ToString();

                                        DateTime dateOf = DateTime.Parse(String.Concat(s, " ", s2));

                                        string fromPhone = row[3].ToString();
                                        string toPhone = row[5].ToString();

                                        string duration = row[9].ToString();
                                        if (duration.Length < 2) duration = "0" + duration;
                                        duration += ":00";

                                        Calls calls = new Calls(dateOf, @fromPhone, @toPhone, duration);
                                        _context.Calls.Add(calls);
                                    }
                                }
                                //excelReader.Close();

                                //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                //process.Start();
                                //process.Close();

                                reader.Close();
                                oledbConn.Close();

                                filesForDelete = file.Substring(0, file.IndexOf('.'));

                                info.Arguments = "/C del /f \"" + filesForDelete + '*'+"\"";
                                process.Start();
                                process.WaitForExit();
                                process.Close();

                                break;

                            case "ЛС_№642000094108":
                                for (int i = 0; i <= 4; i++)
                                {
                                    //dt.Rows[i].Delete();
                                    table.Rows[i].Delete();
                                }

                                //foreach (DataRow row in dt.Rows)
                                foreach (DataRow row in table.Rows)
                                {
                                    if (row.RowState != DataRowState.Deleted && !row[3].ToString().Contains("трафик"))
                                    {
                                        string s = row[2].ToString();
                                        DateTime dateOf = DateTime.Parse(s);
                                        string fromPhone = row[1].ToString();
                                        string toPhone = row[5].ToString();
                                        string duration;
                                        if (row[6].ToString() == "сек")
                                        {
                                            int totalSeconds = Convert.ToInt16(row[7]);
                                            int seconds = totalSeconds % 60;
                                            int minute = (totalSeconds - seconds) / 60;

                                            duration = minute.ToString();
                                            if (duration.Length < 2) duration = "0" + duration;
                                            string durationSeconds = seconds.ToString();
                                            if (durationSeconds.Length < 2) durationSeconds = "0" + durationSeconds;
                                            duration += ":" + durationSeconds;
                                        }
                                        else
                                        {
                                            duration = row[7].ToString();
                                            if (duration.Length < 2) duration = "0" + duration;
                                            duration += ":00";
                                        }

                                        Calls calls = new Calls(dateOf, fromPhone, toPhone, duration);
                                        _context.Calls.Add(calls);
                                    }
                                }
                                //excelReader.Close();

                                //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                //process.Start();
                                //process.Close();

                                reader.Close();
                                oledbConn.Close();

                                filesForDelete = file.Substring(0, file.IndexOf('.'));

                                info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                process.Start();
                                process.WaitForExit();
                                process.Close();

                                break;

                            case "МастерТел":

                                //dt.Rows[0].Delete();
                                table.Rows[0].Delete();
                                //dt.Rows[1].Delete();
                                table.Rows[1].Delete();

                                //foreach (DataRow row in dt.Rows)
                                foreach (DataRow row in table.Rows)
                                {
                                    if (row.RowState != DataRowState.Deleted)
                                    {
                                        string s = row[5].ToString().Substring(0, 10);
                                        string s2 = DateTime.Parse(row[6].ToString()).ToString().Substring(11);
                                        DateTime dateOf = DateTime.Parse(String.Concat(s, " ", s2));

                                        string fromPhone = row[0].ToString();
                                        string toPhone = row[1].ToString();

                                        string duration = row[7].ToString();
                                        if (duration.Length < 2) duration = "0" + duration;
                                        duration += ":00";

                                        Calls calls = new Calls(dateOf, @fromPhone, @toPhone, duration);
                                        _context.Calls.Add(calls);
                                    }

                                    //if (row.RowState != DataRowState.Deleted)
                                    {
                                        //string s = row[0].ToString();
                                        //s = s.Replace("\"", "");
                                        //string s1 = s.Substring(0, s.IndexOf(';'));
                                        //string callFrom = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //string callTo = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //s = s.Substring(s.IndexOf(';') + 1);
                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //s = s.Substring(s.IndexOf(';') + 1);
                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //s = s.Substring(s.IndexOf(';') + 1);
                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //string dOf = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //dOf += " " + s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //string durationOf = s1;

                                        //DateTime dateOf = DateTime.Parse(dOf);
                                        //string fromPhone = callFrom;
                                        //string toPhone = callTo;

                                        //string duration = durationOf;
                                        //if (duration.Length < 2) duration = "0" + duration;
                                        //duration += ":00";

                                        //Calls calls = new Calls(dateOf, fromPhone, toPhone, duration);
                                        //_context.Calls.Add(calls);
                                    }
                                }
                                //excelReader.Close();

                                //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                //process.Start();
                                //process.Close();

                                reader.Close();
                                oledbConn.Close();

                                filesForDelete = file.Substring(0, file.IndexOf('.'));

                                info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                process.Start();
                                process.WaitForExit();
                                process.Close();

                                break;

                            case "МУС":

                                //dt.Rows[0].Delete();
                                table.Rows[0].Delete();

                                //foreach (DataRow row in dt.Rows)
                                foreach (DataRow row in table.Rows)
                                {
                                    if (row.RowState != DataRowState.Deleted)
                                    {
                                        //string s = row[5].ToString().Substring(0, 10);
                                        //string s2 = DateTime.Parse(row[6].ToString()).ToString().Substring(11);
                                        //DateTime dateOf = DateTime.Parse(String.Concat(s, " ", s2));
                                        DateTime dateOf = DateTime.Parse(row[0].ToString());

                                        string fromPhone = row[2].ToString();
                                        string toPhone = row[1].ToString();

                                        string durationOf = row[3].ToString();
                                        string durationSeconds = durationOf.Substring(durationOf.IndexOf(',') + 1);
                                        int seconds = Int16.Parse(durationSeconds) / 5 * 3;
                                        durationSeconds = seconds.ToString();

                                        if (durationSeconds.Length < 2) durationSeconds = "0" + durationSeconds;

                                        string durationMinutes;
                                        if (durationOf.Length > 1)
                                            durationMinutes = durationOf.Substring(0, durationOf.IndexOf(","));
                                        else durationMinutes = durationOf;
                                        
                                        if (durationMinutes.Length < 2) durationMinutes = "0" + durationMinutes;
                                        durationOf = durationMinutes + ":" + durationSeconds;

                                        //string s = row[0].ToString();
                                        //s = s.Replace("\"", "");
                                        //string s1 = s.Substring(0, s.IndexOf(';'));
                                        //string dOf = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //string callFrom = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //s1 = s.Substring(0, s.IndexOf(';'));
                                        //string callTo = s1;
                                        //s = s.Substring(s.IndexOf(';') + 1);

                                        //string durationOf = s;
                                        //if (durationOf.Length < 2) durationOf = "0" + durationOf;
                                        //int seconds = Int16.Parse(row[1].ToString().Substring(0, 2));
                                        //string durationSeconds = (seconds / 5 * 3).ToString();
                                        //if (durationSeconds.Length < 2) durationSeconds = "0" + durationSeconds;
                                        //durationOf += ":" + durationSeconds.ToString();

                                        //DateTime dateOf = DateTime.Parse(dOf);
                                        //string fromPhone = callFrom;
                                        //string toPhone = callTo;
                                        //string duration = durationOf;

                                        Calls calls = new Calls(dateOf, fromPhone, toPhone, durationOf);
                                        _context.Calls.Add(calls);
                                    }
                                }
                                //excelReader.Close();

                                //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                //process.Start();
                                //process.Close();

                                reader.Close();
                                oledbConn.Close();

                                filesForDelete = file.Substring(0, file.IndexOf('.'));

                                info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                process.Start();
                                process.WaitForExit();
                                process.Close();

                                break;

                            case "ОБИТ":
                                if (newFile.Contains("входящие"))
                                {
                                    //foreach (DataRow row in dt.Rows)
                                    foreach (DataRow row in table.Rows)
                                    {
                                        string ss = row[0].ToString();
                                        string s = row[0].ToString().Substring(0, 19);
                                        DateTime dateOf = DateTime.Parse(String.Concat(s));
                                        string fromPhone = row[0].ToString().Substring(20, 10);
                                        //string toPhone = row[0].ToString().Substring(30, row[0].ToString().Length - 30);
                                        //string duration = row[0].ToString().Substring(39, 1);

                                        Calls calls = new Calls(dateOf, fromPhone, "Либхерр Русланд", "");
                                        _context.Calls.Add(calls);
                                    }
                                    //excelReader.Close();

                                    //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                    //process.Start();
                                    //process.Close();

                                    reader.Close();
                                    oledbConn.Close();

                                    filesForDelete = file.Substring(0, file.IndexOf('.'));

                                    info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                    process.Start();
                                    process.WaitForExit();
                                    process.Close();
                                }
                                else
                                {
                                    //foreach (DataRow row in dt.Rows)
                                    foreach (DataRow row in table.Rows)
                                    {
                                        string s = row[0].ToString();
                                        bool b = true;
                                        int i = 1;
                                        while (b)
                                        {
                                            try
                                            {
                                                s += row[i].ToString();
                                                i++;
                                            }
                                            catch
                                            {
                                                break;
                                            }
                                        }

                                        s = s.Replace("?", "");
                                        s = s.Replace("L", "");
                                        s = s.Replace("D", "");
                                        s = s.Replace("mob", "");
                                        s = s.Replace(",", "");
                                        s = s.Replace("(", "");
                                        s = s.Replace(")", "");
                                        s = s.Replace("EF", "");
                                        s = s.Replace(" ", "");
                                        s = s.Replace("--", " ");
                                        s = s.Replace("\t", "");

                                        string s1 = s.Substring(0, 10) + " " + s.Substring(10, 8);
                                        DateTime dateOf = DateTime.Parse(String.Concat(s1));
                                        //string fromPhone = s.Substring(18, 10);
                                        string fromPhone = "Либхерр Русланд";
                                        string toPhone = s.Substring(29, 14).Replace("-", "");

                                        s = s.Substring(43);
                                        string duration = s.ToString().Substring(0, s.Length - s.IndexOf('.')).Replace(".", "");

                                        int totalSeconds = Convert.ToInt16(duration);
                                        int seconds = totalSeconds % 60;
                                        int minute = (totalSeconds - seconds) / 60;

                                        duration = minute.ToString();
                                        if (duration.Length < 2) duration = "0" + duration;
                                        string durationSeconds = seconds.ToString();
                                        if (durationSeconds.Length < 2) durationSeconds = "0" + durationSeconds;
                                        duration += ":" + durationSeconds;

                                        Calls calls = new Calls(dateOf, fromPhone, toPhone, duration);
                                        _context.Calls.Add(calls);
                                    }
                                    //excelReader.Close();

                                    //process.StartInfo.Arguments = "/c del \"" + newFile + "\" /s";
                                    //process.Start();
                                    //process.Close();

                                    reader.Close();
                                    oledbConn.Close();

                                    filesForDelete = file.Substring(0, file.IndexOf('.'));

                                    info.Arguments = "/C del /f \"" + filesForDelete + '*';
                                    process.Start();
                                    process.WaitForExit();
                                    process.Close();
                                }
                                break;

                        }
                        //excelReader.Close();
                        //stream.Close();
                    }
                }

            }

            _context.SaveChanges();
            //x1.Quit();
        }
    }
}