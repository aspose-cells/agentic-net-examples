// Title: How to show and print gridlines on an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable gridlines for printed worksheet | Print Excel sheet with visible gridlines using Aspose.Cells .NET | Set PageSetup.PrintGridlines and IsGridlinesVisible in Aspose.Cells example | Export workbook to PDF with gridlines using Aspose.Cells for .NET | Aspose.Cells keep gridlines visible when saving as PDF
// Tags: Aspose.Cells PrintGridlines API | worksheet IsGridlinesVisible property C# | export workbook to PDF with gridlines Aspose.Cells | page setup print options Aspose.Cells .NET | show gridlines in printed Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace GridlinesPrintDemo
{
    // The sample loads an existing workbook or creates a new one, sets the worksheet's IsGridlinesVisible flag to true, enables PageSetup.PrintGridlines, configures additional print settings, and saves the file as both XLSX and PDF to demonstrate gridlines in the UI and in the printed/PDF output.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input workbook path (optional)
                string inputPath = args.Length > 0 ? args[0] : string.Empty;
                Workbook workbook;

                if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Create a new workbook with sample data
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "SampleData";

                    // Populate some cells so gridlines are visible
                    ws.Cells["A1"].PutValue("Item");
                    ws.Cells["B1"].PutValue("Quantity");
                    ws.Cells["A2"].PutValue("Apples");
                    ws.Cells["B2"].PutValue(10);
                    ws.Cells["A3"].PutValue("Oranges");
                    ws.Cells["B3"].PutValue(15);
                    ws.AutoFitColumns();

                    Console.WriteLine("Created a new workbook with sample data.");
                }

                // Get the first worksheet (or the active one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure gridlines are visible on the sheet itself
                worksheet.IsGridlinesVisible = true;

                // Enable gridlines in the printed output
                worksheet.PageSetup.PrintGridlines = true;

                // Optional: set other print options for clarity
                worksheet.PageSetup.PrintHeadings = false; // hide row/column headings
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0; // let height adjust automatically

                // Define output file names
                string outputExcel = "GridlinesEnabled.xlsx";
                string outputPdf = "GridlinesEnabled.pdf";

                // Save the workbook as Excel (to verify the setting in the UI)
                workbook.Save(outputExcel, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with gridlines visible: {outputExcel}");

                // Save as PDF to demonstrate printed appearance with gridlines
                workbook.Save(outputPdf, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved as PDF with gridlines printed: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
