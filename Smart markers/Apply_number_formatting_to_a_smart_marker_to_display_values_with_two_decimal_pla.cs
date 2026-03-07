using System;
using System.Collections;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerNumberFormatDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Value");

            // Insert smart markers for data rows (B2:B5 will hold the numeric values)
            for (int row = 2; row <= 5; row++)
            {
                sheet.Cells[row - 1, 0].PutValue("&=Data.Name");   // Column A
                sheet.Cells[row - 1, 1].PutValue("&=Data.Value"); // Column B
            }

            // ------------------------------------------------------------
            // Apply a number format ("0.00") to the cells that hold the smart markers
            // ------------------------------------------------------------
            Style numberStyle = workbook.CreateStyle();
            numberStyle.Number = 2; // 0.00 format

            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            AsposeRange smartMarkerRange = sheet.Cells.CreateRange("B2:B5");
            smartMarkerRange.ApplyStyle(numberStyle, flag);

            // ------------------------------------------------------------
            // Prepare data source and process the smart markers
            // ------------------------------------------------------------
            ArrayList data = new ArrayList();
            data.Add(new { Name = "Product A", Value = 123.456 });
            data.Add(new { Name = "Product B", Value = 78.9 });
            data.Add(new { Name = "Product C", Value = 45.0 });
            data.Add(new { Name = "Product D", Value = 0.12 });

            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process();

            // ------------------------------------------------------------
            // Save the resulting workbook
            // ------------------------------------------------------------
            workbook.Save("SmartMarkerNumberFormatted.xlsx");
        }
    }
}