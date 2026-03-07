using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(3);

            // Set the print area to the range A1:B3
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true
            };

            // Save the workbook; only the print area will be included in the output file
            workbook.Save("PrintAreaDemo.html", options);
        }
    }
}