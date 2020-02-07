using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Excel_Create
{
    class Excel
    {
        public Application excel = new Microsoft.Office.Interop.Excel.Application();
        public Workbook workbook;
        public Worksheet worksheet;

        public Excel(string path, string workSheet)
        {
            workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            worksheet = workbook.Worksheets[1];
            worksheet.Name = workSheet;

            worksheet.Cells[1, 1] = "People";
            worksheet.get_Range("A1").Interior.ColorIndex = 4;
            worksheet.get_Range("A1", "D1").Merge();
            worksheet.get_Range("A1").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.get_Range("A1").Font.Size = 25;

            worksheet.Cells[2, 1] = "Id";
            //worksheet.get_Range("A2").Orientation = 90;
            worksheet.get_Range("A2").Interior.ColorIndex = 6;
            worksheet.get_Range("A2", "A1000").HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[2, 2] = "First Name";
            //worksheet.get_Range("B2").Orientation = 90;
            worksheet.get_Range("B2").Interior.ColorIndex = 6;
            worksheet.get_Range("B2", "B1000").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.get_Range("B2").EntireColumn.ColumnWidth = 50;

            worksheet.Cells[2, 3] = "Last Name";
            //worksheet.get_Range("C2").Orientation = 90;
            worksheet.get_Range("C2").Interior.ColorIndex = 6;
            worksheet.get_Range("C2").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.get_Range("C2").EntireColumn.ColumnWidth = 50;

            worksheet.Cells[2, 3] = "Last Name";
            //worksheet.get_Range("C2").Orientation = 90;
            worksheet.get_Range("C2").Interior.ColorIndex = 6;
            worksheet.get_Range("C2", "C1000").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.get_Range("C2").EntireColumn.ColumnWidth = 50;

            worksheet.Cells[2, 4] = "Age";
            //worksheet.get_Range("C2").Orientation = 90;
            worksheet.get_Range("D2").Interior.ColorIndex = 6;
            worksheet.get_Range("D2", "D1000").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.get_Range("D2").EntireColumn.ColumnWidth = 20;

            workbook.SaveAs(path);
            excel.Visible = true;
            excel.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            
            //workbook.Close();
            //excel.Quit();
        }

        public void ExcelCreateTableBody(List<Person> people)
        {
            for (int i = 3; i < people.Count + 3; i++)
            {
                worksheet.Cells[i, 1] = people[i - 3].Id;
                worksheet.Cells[i, 2] = people[i - 3].FirstName;
                worksheet.Cells[i, 3] = people[i - 3].LastName;
                worksheet.Cells[i, 4] = people[i - 3].Age;
            }
            Shape shape = worksheet.Shapes.AddChart(XlChartType.xlLine, 400, 5, 1000, 200);
            Chart chart = shape.Chart;

            Range range = worksheet.get_Range("C2", "D104");
            Series series = (Series)chart.SeriesCollection(1);
            series.XValues = range;
        }
    }
}
