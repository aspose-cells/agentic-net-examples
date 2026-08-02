// Title: Aspose.Cells for .NET – Repeat Rows 1‑2 as Print Titles on Every Page
// Description: The C# sample builds a new workbook, selects the first sheet, assigns "$1:$2" to PageSetup.PrintTitleRows so that rows 1‑2 are printed at the top of each page, and writes the result to PrintTitleRows_1_2.xlsx.
// Keywords: Aspose.Cells PrintTitleRows | C# repeat header rows printing | Aspose.Cells PageSetup example | set worksheet print titles .NET | repeat rows on each printed page
// Common Searches: how to repeat header rows with Aspose.Cells .NET | C# set print title rows Aspose.Cells | Aspose.Cells repeat rows on every page | PageSetup.PrintTitleRows usage example
// Developer Intent: Configure a worksheet so specific rows are printed as titles on every page.
// Use Cases: Printing large reports where the first two rows contain column headings. | Creating multi‑page invoices that need a static title block. | Generating data sheets with fixed header rows for consistent printouts.
// AI Prompts: Show me C# code to set rows 1‑2 as repeating print titles using Aspose.Cells. | Give a complete Aspose.Cells example that saves a workbook after defining print title rows. | Explain how to change or clear the PrintTitleRows setting in an existing worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The C# sample builds a new workbook, selects the first sheet, assigns "$1:$2" to PageSetup.PrintTitleRows so that rows 1‑2 are printed at the top of each page, and writes the result to PrintTitleRows_1_2.xlsx.
    public class SetPrintTitleRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Repeat rows 1 through 2 at the top of each printed page
                worksheet.PageSetup.PrintTitleRows = "$1:$2";

                // Save the workbook
                workbook.Save("PrintTitleRows_1_2.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrintTitleRowsDemo.Run();
        }
    }
}
