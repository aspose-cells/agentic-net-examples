using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(699.99);

            // Create a style to apply
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.White;
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Define which style attributes to apply
            StyleFlag flag = new StyleFlag();
            flag.All = true; // apply all style attributes

            // Apply the style to the header range A1:B1
            var headerRange = worksheet.Cells.CreateRange("A1:B1");
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook to an XLSX file
            workbook.Save("ExportedData.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook exported successfully with values and formatting.");
        }
    }
}