using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FileUploadAndSaveToSQL.Controllers
{
    public class HomeController : Controller
    {
        SqlConnection Connection = new SqlConnection("Data Source=DESKTOP-TAS8QG0;Initial Catalog=test;Integrated Security=True");
        SqlCommand Command = new SqlCommand();

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Save(HttpPostedFileBase excelFile)
        {
            if (excelFile == null || excelFile.ContentLength==0)
            {
                ViewBag.error = "please select excel file <br />";
                return View("Index");
            }
            else
            {
                if (excelFile.FileName.EndsWith("xls") || excelFile.FileName.EndsWith("xlsx"))
                {
                    
                    return saveToDatabase(excelFile);

                }
                else
                {
                    ViewBag.error = "file type is incorrect <br />";
                    return View("Index");
                }
            }
        }
        public ViewResult saveToDatabase (HttpPostedFileBase excelFile)
        {
            try
            {
                //Save the uploaded Excel file.
                string path = Server.MapPath("~/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                excelFile.SaveAs(path + Path.GetFileName(excelFile.FileName));


                using (XLWorkbook workBook = new XLWorkbook(path + Path.GetFileName(excelFile.FileName)))
                {
                    String columnsNamesQuery = String.Format("CREATE TABLE {0} (", excelFile.FileName.Split('.')[0]);
                    String insertQuery = String.Format("INSERT INTO {0} (", excelFile.FileName.Split('.')[0]);

                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                columnsNamesQuery += "[" + cell.Value.ToString() + "] VARCHAR(MAX),";
                                insertQuery += "[" + cell.Value.ToString() + "],";
                            }

                            insertQuery = insertQuery.Remove(insertQuery.Length - 1, 1);
                            insertQuery += ") VALUES (";
                            columnsNamesQuery = columnsNamesQuery.Remove(columnsNamesQuery.Length - 1, 1);
                            columnsNamesQuery += ");";

                            //Creating data table
                            Command.Connection = Connection;
                            Connection.Open();
                            Command.CommandText = columnsNamesQuery;
                            Command.ExecuteNonQuery();
                            Connection.Close();

                            firstRow = false;
                        }
                        else
                        {
                            String rowsQuery = insertQuery;
                            foreach (IXLCell cell in row.Cells(workSheet.FirstCellUsed().Address.ColumnNumber, workSheet.LastCellUsed().Address.ColumnNumber))
                            {
                                if (cell.IsEmpty())
                                {
                                    rowsQuery += "NULL,";
                                }
                                else
                                {
                                    rowsQuery += "'" + cell.Value.ToString() + "' ,";
                                }
                            }
                            rowsQuery = rowsQuery.Remove(rowsQuery.Length - 1, 1);
                            rowsQuery += ");";

                            //Insert data into database
                            Command.Connection = Connection;
                            Connection.Open();
                            Command.CommandText = rowsQuery;
                            Command.ExecuteNonQuery();
                            Connection.Close();
                        }

                    }
                }
                return View("Success");
            }
            catch (Exception e)
            {
                ViewBag.error = "ERROR: " + e.Message + "<br />";
                return View("Index");
            }
        }
    }    
}