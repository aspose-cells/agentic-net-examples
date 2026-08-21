// Title: Print Gridlines and Row/Column Headings with Aspose.Cells for .NET
// Description: Shows how to enable PrintGridlines and PrintHeadings via Worksheet.PageSetup, optionally define a PrintArea, and save the workbook so the printed page mirrors the on‑screen view.
// Keywords: Aspose.Cells | .NET | C# | PrintGridlines | PrintHeadings | PageSetup | print area | gridlines on print | row headings | column headings | Excel print settings
// Common Searches: Aspose.Cells print gridlines and headings | C# enable PrintHeadings in Excel export | set PrintArea with Aspose.Cells | how to print Excel gridlines using Aspose.Cells .NET | Aspose.Cells page setup print options
// Developer Intent: Configure a worksheet to print both gridlines and row/column headings so the hard‑copy matches the on‑screen layout.
// Use Cases: Generate a sales report that retains gridlines and column headers when printed. | Create printable invoices where row numbers and column letters aid manual reference. | Define a specific print area for a data subset while preserving visual grid structure.
// AI Prompts: Provide C# code that sets PrintGridlines, PrintHeadings, and a custom PrintArea using Aspose.Cells. | Show how to configure PageSetup to print gridlines, headings, and fit the worksheet to one page in Aspose.Cells for .NET. | Explain how to programmatically verify that PrintGridlines and PrintHeadings settings are applied after saving the workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to enable PrintGridlines and PrintHeadings via Worksheet.PageSetup, optionally define a PrintArea, and save the workbook so the printed page mirrors the on‑screen view.
    public class PrintGridlinesAndHeadingsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(2.5);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(1.8);

            // Enable printing of gridlines and row/column headings
            worksheet.PageSetup.PrintGridlines = true;   // print cell gridlines
            worksheet.PageSetup.PrintHeadings = true;   // print row (1,2,…) and column (A,B,…) headings

            // Optionally define a print area to limit what gets printed
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Define output file path
            string outputPath = "PrintGridlinesAndHeadings.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook (Excel format retains the print settings)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
