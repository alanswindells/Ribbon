using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;

namespace Ribbon
{

    public class DataWriter
    {
        public static int NumOccurances(string target, string match)
        {
            int pos = -1;
            int n = 0;

            for (; (pos = target.IndexOf(match, pos + 1)) >= 0; n++) ;

            return n;
        }

        public static void WriteData()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            /*            Worksheet ws = wb.Worksheets.Item["Sheet1"];
                        ws.Activate();
                        ws.Range["A1"].Value = "Date";
                        ws.Range["B1"].Value = "Value";

                        Range headerRow = ws.Range["A1", "B1"];
                        headerRow.Font.Size = 12;
                        headerRow.Font.Bold = true;

                        // Generally it's faster to write an array to a range
                        var values = new object[100, 2];
                        var startDate = new DateTime(2007, 1, 1);
                        var rand = new Random();
                        for (int i = 0; i < 100; i++)
                        {
                            values[i, 0] = startDate.AddDays(i);
                            values[i, 1] = rand.NextDouble();
                        }

                        ws.Range["A2"].Resize[100, 2].Value = values;
                        ws.Columns["A:A"].EntireColumn.AutoFit();

                        // Add a chart
                        Range dataRange= ws.Range["A1:B101"];
                        dataRange.Select();
                        ws.Shapes.AddChart(XlChartType.xlColumnClustered).Select();
                        xlApp.ActiveChart.SetSourceData(Source: dataRange);
            */

            string filename = "D:\\OneDrive\\Alan\\Code\\Visual Studio\\Projects\\Ribbon\\bin\\Debug\\ID1541v1-FUZZY-results-extract.csv";

            System.Windows.Forms.OpenFileDialog fd = new System.Windows.Forms.OpenFileDialog();
            fd.Filter = "CSV Files|*.csv";
            fd.InitialDirectory = wb.Path;

            if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename = fd.FileName;
            }

            xlApp.ScreenUpdating = false;
            
            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Name = "data";
            
            var columns = File.ReadLines("D:\\OneDrive\\Alan\\Code\\Visual Studio\\Projects\\Ribbon\\bin\\Debug\\test.txt").First().Split('\t').ToArray();
            int cols = columns.Length;
            var contents = File.ReadAllLines(filename);
            int rows = contents.Length;
            var headings = contents.First().Split('\t').ToArray();
            var vals = new object[rows, cols + 2];
            var colInd = new int[cols];
            
            for (int i = 0; i < cols; i++) colInd[i] = Array.IndexOf(headings, columns[i]);

            int r = 1;
            int depthCol = Array.IndexOf(headings, "Depth");
            int nameCol = Array.IndexOf(headings, "Test Name");
                
            for (int j = 0; j < columns.Length; j++) vals[0, j] = headings[colInd[j]];
            vals[0, cols] = "Len";
            vals[0, cols + 1] = "Words";

            for (int i = 1; i < contents.Length; i++)
            {
                var thisLine = contents[i].Split('\t').ToArray();

                if (Int32.Parse(thisLine[depthCol]) == 1)
                {
                    for (int j = 0; j < Math.Min(thisLine.Length, cols); j++)
                        vals[r, j] = thisLine[colInd[j]];
                    vals[r, cols] = thisLine[nameCol].Length;
                    vals[r, cols + 1] = NumOccurances(thisLine[nameCol], " ");
                    r++;
                }
            }
            Range data = ws.Range["A1"].Resize[r, cols + 2];
            data.Value = vals;

            Worksheet ws2 = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            Range pivotDestination = ws2.Range["A1"];
            string pivotTableName = @"mypivot";

            wb.PivotTableWizard(XlPivotTableSourceType.xlDatabase, data, pivotDestination, pivotTableName, true, true, true, true, 
                Type.Missing, Type.Missing, false, false, XlOrder.xlDownThenOver, 0, Type.Missing, Type.Missing);

            PivotTable pivotTable = (PivotTable)ws2.PivotTables(pivotTableName);

            PivotField shortnamePivotField = (PivotField)pivotTable.PivotFields(2);
                //itemcodePivotField = (PivotField)pivotTable.PivotFields(3);
                //descriptionPivotField = (PivotField)pivotTable.PivotFields(4);
                //pricePivotField = (PivotField)pivotTable.PivotFields(7);

                // Format the Pivot Table.
                //pivotTable.Format(XlPivotFormatType.xlReport2);
                //pivotTable.InGridDropZones = false;
                //pivotTable.SmallGrid = false;
                //pivotTable.ShowTableStyleRowStripes = true;
                //pivotTable.TableStyle2 = "PivotStyleLight1";

                // Page Field
                shortnamePivotField.Orientation = XlPivotFieldOrientation.xlPageField;
                /*shortnamePivotField.Position = 1;
                shortnamePivotField.CurrentPage = "(All)";

                // Row Fields
                descriptionPivotField.Orientation = XlPivotFieldOrientation.xlRowField;
                descriptionPivotField.Position = 1;
                itemcodePivotField.Orientation = XlPivotFieldOrientation.xlRowField;
                itemcodePivotField.Position = 2;

                // Data Field
                pricePivotField.Orientation = XlPivotFieldOrientation.xlDataField;
                pricePivotField.Function = XlConsolidationFunction.xlSum;
  */

            xlApp.ScreenUpdating = true;
        }
    }


}
