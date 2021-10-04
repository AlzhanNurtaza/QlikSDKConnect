using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Qlik;
using Qlik.Engine;
using Qlik.Engine.Communication;
using Qlik.Sense.Client;
using Newtonsoft.Json;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace QlikSDKConnect
{
    class Program
    {
        public class ColumnList
        {
            public string ColumnName { get; set; }
            public Type ColumnType { get; set; }
        }
        public class FilterBlock
        {
            public IList<FilterField> Fields { get; set; }
        }
        public class FilterField
        {
            public string Name { get; set; }
            public IList<FilterValue> FilterValues { get; set; }
        }
        public class FilterValue
        {
            public string Value { get; set; }
        }
        static void Main(string[] args)
        {
            try
            {
                //Loading main conf info from app.config
                Console.OutputEncoding = Encoding.UTF8;
                string qsUri = ConfigurationManager.AppSettings["QlikSenseUri"];
                string qsCertPath = ConfigurationManager.AppSettings["QlikSenseCertPath"];
                string qsDomen = ConfigurationManager.AppSettings["QlikSenseDomen"];
                string qsUser = ConfigurationManager.AppSettings["QlikSenseUser"];
                string qsAppId = ConfigurationManager.AppSettings["QlikSenseAppId"];
                string qsObjectId = ConfigurationManager.AppSettings["QlikSenseObjectId"];
                string qsBookmarkId = ConfigurationManager.AppSettings["QlikSenseBookmarkId"];
                string qsFilters = ConfigurationManager.AppSettings["QlikSenseFilters"];
                string excelSavePath = ConfigurationManager.AppSettings["ExcelSavePath"];


                // The default port number is 4747 but can be customized
                var uri = new Uri(qsUri);
                var certs = CertificateManager.LoadCertificateFromDirectory(qsCertPath);

                ILocation location = Qlik.Engine.Location.FromUri(uri);
                location.AsDirectConnection(qsDomen, qsUser, certificateCollection: certs);


                AppIdentifier appIdentifier = (AppIdentifier)location.AppWithId(qsAppId);
                using (var app = location.App(appIdentifier))
                {



                    if (!string.IsNullOrEmpty(qsBookmarkId))
                        app.ApplyBookmark(qsBookmarkId);


                    if (!string.IsNullOrEmpty(qsFilters))
                    {
                        FilterBlock filterObject = new FilterBlock();
                        List<FilterField> fFields = new List<FilterField>();

                        string[] filterBlocks = qsFilters.Split(';').Select(s => s.Trim()).ToArray();
                        foreach (string block in filterBlocks)
                        {
                            string fName = block.Split(':')[0];
                            FilterField fField = new FilterField();
                            fField.Name = fName;



                            string filterValuePart = block.Split(':')[1];
                            string[] FilterValues = filterValuePart.Split(',');
                            List<FilterValue> filterValueList = new List<FilterValue>();
                            foreach (string value in FilterValues)
                            {
                                FilterValue fv = new FilterValue();
                                fv.Value = value;
                                filterValueList.Add(fv);
                            }
                            fField.FilterValues = (filterValueList);
                            fFields.Add(fField);

                        }
                        filterObject.Fields = fFields;

                        foreach (FilterField field in filterObject.Fields)
                        {
                            var fieldLst = new List<FieldValue>();

                            foreach (FilterValue ff in field.FilterValues)
                            {
                                FieldValue fieldValue = new FieldValue();
                                fieldValue.IsNumeric = false;
                                fieldValue.Text = ff.Value;
                                fieldLst.Add(fieldValue);
                            }

                            app.GetField(field.Name).SelectValues(fieldLst, true);
                        }
                    }





                    var qobject = app.GetGenericObject(qsObjectId);

                    GenericObjectLayout qlayout = qobject.GetLayout();


                    dynamic jsonLayout = JsonConvert.DeserializeObject(qlayout.ToString());
                    //Get header from  qDimensionInfo и qMeasureInfo
                    List<ColumnList> columnList = new List<ColumnList>();
                    foreach (var qDim in jsonLayout.qHyperCube.qDimensionInfo)
                    {
                        ColumnList cl = new ColumnList();
                        cl.ColumnName = Convert.ToString(qDim.qFallbackTitle);
                        cl.ColumnType = typeof(string);
                        columnList.Add(cl);
                    }
                    foreach (var qMes in jsonLayout.qHyperCube.qMeasureInfo)
                    {
                        ColumnList cl = new ColumnList();
                        cl.ColumnName = Convert.ToString(qMes.qFallbackTitle);
                        cl.ColumnType = typeof(double);
                        columnList.Add(cl);
                    }


                    //Get column order of HyperCube 
                    List<int> qColumnOrder = jsonLayout.qHyperCube.qColumnOrder.ToObject<List<int>>();
                    List<string> dtColumnList = qColumnOrder.Select(i => columnList[i].ColumnName).ToList();
                    List<Type> dtColumnListType = qColumnOrder.Select(i => columnList[i].ColumnType).ToList();

                    //Now creating DataTable and creating columns
                    System.Data.DataTable dt = new System.Data.DataTable();
                    for (int i = 0; i < dtColumnList.Count; i++)
                    {
                        dt.Columns.Add(dtColumnList[i], dtColumnListType[i]);
                    }


                    var pager = qobject.GetAllHyperCubePagers().First();
                    var allPages = pager.IteratePages(new[] { new NxPage { Width = pager.NumberOfColumns, Height = 100 } }, Pager.Next).Select(pageSet => pageSet.Single());
                    var datas = allPages.SelectMany(page => page.Matrix);

                    foreach (var data in datas)
                    {
                        System.Data.DataRow row = dt.NewRow();
                        int columnIndex = 0;
                        foreach (var column in dt.Columns)
                        {
                            if (data[columnIndex].State.ToString() == "LOCKED")
                            {
                                if (data[columnIndex].Num.ToString() != "NaN")
                                {
                                    row[columnIndex] = data[columnIndex].Num;
                                }
                                else
                                {
                                    row[columnIndex] = DBNull.Value;
                                }

                            }
                            else
                            {
                                row[columnIndex] = data[columnIndex].Text;
                            }
                            columnIndex++;
                        }
                        dt.Rows.Add(row);

                    }

                    ExportToExcel(dt, excelSavePath);

                }
            }
            catch(Exception ex)
            {
                ErrorLogging(ex);
            }
            

        }
        public static void ErrorLogging(Exception ex)
        {
            string strPath = @"C:\QlikSDKConnect\Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);

            }
        }

        static void ExportToExcel(System.Data.DataTable tbl, string excelFilePath = null)
        {
            // load excel, and create a new workbook
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            var workBooks = excelApp.Workbooks;
            var workBook = workBooks.Add();
            // single worksheet
            Excel._Worksheet workSheet = workBook.ActiveSheet;
            workSheet.Name = "output";

            try
            {
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");



                // column headings
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = tbl.Columns[i].ColumnName;
                }

                // rows
                for (var i = 0; i < tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1].NumberFormat = tbl.Columns[j].DataType == typeof(string) ? "@" : "0.00";
                        workSheet.Cells[i + 2, j + 1] = tbl.Rows[i][j];
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    workBook.SaveCopyAs(@excelFilePath);
                    workBooks.Close();
                    excelApp.Quit();
                    if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                    if (workBooks != null) Marshal.ReleaseComObject(workBooks);
                    if (workBook != null) Marshal.ReleaseComObject(workBook);
                    if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                    workBooks = null;
                    workBook = null;
                    excelApp = null;
                    GC.Collect();

                }
            }
            catch (Exception ex)
            {
                ErrorLogging(ex);
                workBooks.Close();
                excelApp.Quit();
                if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                if (workBooks != null) Marshal.ReleaseComObject(workBooks);
                if (workBook != null) Marshal.ReleaseComObject(workBook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                workBooks = null;
                workBook = null;
                excelApp = null;
                GC.Collect();
            }
        }
    }

}
