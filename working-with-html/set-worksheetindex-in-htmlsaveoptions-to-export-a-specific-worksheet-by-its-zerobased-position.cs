// Title: Export a specific worksheet to HTML by setting ActiveSheetIndex and using HtmlSaveOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export only the second worksheet to HTML | how to use HtmlSaveOptions ExportActiveWorksheetOnly with worksheet index | C# set active sheet index before saving workbook as HTML Aspose.Cells | export specific worksheet by zero‑based index to HTML using Aspose.Cells | save one sheet as HTML file Aspose.Cells .NET example
// Tags: HtmlSaveOptions active worksheet export Aspose.Cells | ActiveSheetIndex workbook property Aspose.Cells | export single worksheet to HTML C# | zero‑based worksheet index selection Aspose.Cells | Aspose.Cells specific sheet HTML conversion

using System;
using Aspose.Cells;

namespace ExportSpecificWorksheetToHtml
{
    // The example creates a workbook with three worksheets, sets the workbook's ActiveSheetIndex to 1 (the second sheet), configures HtmlSaveOptions with ExportActiveWorksheetOnly = true, and saves only that worksheet as an HTML file named 'SecondSheetExport.html'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and add sample data to multiple sheets
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Worksheet 0 (default)
                Worksheet sheet0 = workbook.Worksheets[0];
                sheet0.Name = "FirstSheet";
                sheet0.Cells["A1"].PutValue("This is the first worksheet.");
                sheet0.Cells["A2"].PutValue(123);
                sheet0.Cells["B2"].PutValue(456);

                // Worksheet 1 (second sheet)
                Worksheet sheet1 = workbook.Worksheets.Add("SecondSheet");
                sheet1.Cells["A1"].PutValue("This is the second worksheet.");
                sheet1.Cells["A2"].PutValue("Hello");
                sheet1.Cells["B2"].PutValue("World");

                // Worksheet 2 (third sheet)
                Worksheet sheet2 = workbook.Worksheets.Add("ThirdSheet");
                sheet2.Cells["A1"].PutValue("This is the third worksheet.");
                sheet2.Cells["A2"].PutValue(DateTime.Now);
                sheet2.Cells["B2"].PutValue(3.14159);

                // ---------------------------------------------------------------
                // 2. Configure HtmlSaveOptions to export only the second worksheet
                //    (zero‑based index 1)
                // ---------------------------------------------------------------
                // Set the active sheet to the one we want to export
                workbook.Worksheets.ActiveSheetIndex = 1;

                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export only the active worksheet
                    ExportActiveWorksheetOnly = true
                };

                // ---------------------------------------------------------------
                // 3. Save the workbook as HTML using the configured options
                // ---------------------------------------------------------------
                string outputPath = "SecondSheetExport.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Successfully exported worksheet index 1 to HTML file: {outputPath}");
            }
            catch (Exception ex)
            {
                // Report any runtime errors (e.g., licensing, I/O)
                Console.WriteLine("An error occurred while exporting the worksheet to HTML:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
