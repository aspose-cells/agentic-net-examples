// Title: Set landscape page orientation and freeze the first row header in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet to landscape orientation and freeze top row | How to repeat header row on each printed page using Aspose.Cells .NET | C# code to configure page margins and paper size in Aspose.Cells workbook | Freeze panes first row and keep columns visible with Aspose.Cells | Save an Excel file with custom page setup and frozen header using Aspose.Cells
// Tags: landscape orientation page setup Aspose.Cells | freeze first row panes Aspose.Cells C# | repeat print title rows Aspose.Cells | custom paper size margins Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomOrientationAndFreeze
{
    // The example loads an existing workbook or creates a new one, sets the worksheet to landscape orientation with A4 paper size and custom margins, freezes the first row while keeping all columns scrollable, repeats that header row on every printed page, and saves the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input file (optional) and output file
            string inputPath = args.Length > 0 ? args[0] : string.Empty;
            string outputPath = "CustomOrientationAndFreeze.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if a valid path is supplied
                if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook with sample data
                    Console.WriteLine("No valid input file supplied. Creating a new workbook with sample data.");
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "DataSheet";

                    // Header row
                    ws.Cells["A1"].PutValue("ID");
                    ws.Cells["B1"].PutValue("Name");
                    ws.Cells["C1"].PutValue("Quantity");
                    ws.Cells["D1"].PutValue("Price");

                    // Sample rows
                    for (int i = 2; i <= 20; i++)
                    {
                        ws.Cells[$"A{i}"].PutValue(i - 1);
                        ws.Cells[$"B{i}"].PutValue($"Item {i - 1}");
                        ws.Cells[$"C{i}"].PutValue((i - 1) * 10);
                        ws.Cells[$"D{i}"].PutValue(((i - 1) * 1.5).ToString("F2"));
                    }
                }

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Set custom page orientation (Landscape)
                // -------------------------------------------------
                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
                Console.WriteLine("Page orientation set to Landscape.");

                // Optional: set paper size and margins
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.LeftMargin = 0.5;
                worksheet.PageSetup.RightMargin = 0.5;
                worksheet.PageSetup.TopMargin = 0.75;
                worksheet.PageSetup.BottomMargin = 0.75;

                // -------------------------------------------------
                // 2. Freeze header rows (first row)
                // -------------------------------------------------
                // Freeze the first row (row index 1) while keeping all columns visible
                worksheet.FreezePanes(1, 0, 1, 0);
                Console.WriteLine("Header row frozen (first row).");

                // Repeat the header row on each printed page
                worksheet.PageSetup.PrintTitleRows = "$1:$1";
                Console.WriteLine("Print title rows set to repeat the header row on each printed page.");

                // -------------------------------------------------
                // 3. Save the workbook
                // -------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Report full error details
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.ToString());
                Environment.Exit(1);
            }
        }
    }
}
