using System;
using Microsoft.AnalysisServices.Tabular;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;

namespace Practicing_TOM
{
    class Program
    {
        static Model model = null;
        static Server server = null;
        static void Main(string[] args)
        {

            connectToServer();
            writeMeasuresToExcel(tableName: "Report Measures", getAllMeasures());
            server.Disconnect();
        }

        static void connectToServer()
        {
            server = new Server();
            server.Connect(@"localhost:51369");

            model = server.Databases[0].Model;
            //model = server.Databases.GetByName("Contoso").Model;
        }

        static List<Measure> getAllMeasures()
        {
            var allTables = new List<Table>();
            foreach (Table table in model.Tables) 
            { 
                allTables.Add(table); 
            }

            var allMeasures = new List<Measure>();
            foreach (Table table in allTables)
            {
                foreach (Measure measure in table.Measures) 
                { 
                    allMeasures.Add(measure); 
                }
            }

            return allMeasures;
        }

        static void writeMeasuresToExcel(string tableName, List<Measure> measureList )
        {
            Excel.Application xlApp = new Excel.Application();
            var workboks = xlApp.Workbooks;
            Excel.Workbook wb = workboks.Add();
            Excel.Worksheet ws = wb.Worksheets[1];
            ws.Name = "Measures";

            int rowNumber = 1;
            int colNumber = 0;

            string[] columnHeaders = 
            { 
                "Measure Name", 
                "Expression", 
                "Format String", 
                "Hidden", 
                "Data Type" 
            };

            foreach(Excel.Range r in ws.Range["A1:E1"])
            {
                r.Value = columnHeaders[colNumber];
                colNumber++;
            }

            foreach (Measure m in measureList)
            {
                ws.Range[$"A{rowNumber}"].Value = m.Name;
                ws.Range[$"B{rowNumber}"].Value = m.Expression;
                ws.Range[$"C{rowNumber}"].Value = m.FormatString;
                ws.Range[$"D{rowNumber}"].Value = m.IsHidden;
                ws.Range[$"E{rowNumber}"].Value = m.DataType;
                rowNumber++;
            }

            ws.Columns["B"].WrapText = false;
            ws.Columns.AutoFit();
            ws.Columns["B"].ColumnWidth = 40;

            Excel.ListObject measureTable = ws.ListObjects.AddEx(
                SourceType: Excel.XlListObjectSourceType.xlSrcRange,
                Source: ws.Range[$"A1:E{rowNumber - 1}"],
                LinkSource: Type.Missing,
                XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes
            );
            measureTable.Name = "Measure Table";
            
            string filePath = @"C:\Users\antsharma\Downloads\Power BI Measures.xlsx";

            if (File.Exists(filePath)){ File.Delete(filePath); }

            wb.SaveAs2(filePath);
            wb.Close();
            xlApp.Quit();
        }
    }
}
