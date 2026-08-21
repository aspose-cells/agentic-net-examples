// Title: Enable Black‑and‑White Printing for All Worksheets in Aspose.Cells (C#)
// Description: Creates a workbook, iterates through each worksheet, sets PageSetup.BlackAndWhite to true, and saves the file, allowing the entire workbook to print in monochrome to conserve ink.
// Keywords: Aspose.Cells | C# | BlackAndWhite property | PageSetup | monochrome printing | reduce ink usage | set black and white printing | worksheet page setup
// Common Searches: Aspose.Cells set black and white printing for all sheets | C# enable monochrome printing in Excel workbook using Aspose.Cells | How to reduce ink consumption with Aspose.Cells | PageSetup BlackAndWhite example .NET | Print Excel file in black and white with Aspose.Cells
// Developer Intent: Apply the BlackAndWhite page‑setup option to every worksheet so the workbook prints in monochrome, saving ink.
// Use Cases: Generate printer‑friendly reports that use less ink by enabling black‑and‑white printing on all worksheets before distribution. | Batch‑process existing workbooks to convert them to monochrome for archival or mass mailing. | Automatically enforce monochrome output when exporting financial statements or dashboards to ensure consistent printing across all sheets.
// AI Prompts: Write C# code with Aspose.Cells that sets BlackAndWhite = true for each worksheet and then saves the workbook as a PDF. | Provide an example that toggles the BlackAndWhite setting based on a configuration flag in a .NET application. | Explain how to programmatically verify that the BlackAndWhite property has been applied to all worksheets after saving the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, iterates through each worksheet, sets PageSetup.BlackAndWhite to true, and saves the file, allowing the entire workbook to print in monochrome to conserve ink.
    public class EnableBlackAndWhitePrinting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (empty with a default worksheet)
                Workbook workbook = new Workbook();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the PageSetup object of the worksheet
                    PageSetup pageSetup = sheet.PageSetup;

                    // Enable black‑and‑white printing for this worksheet
                    pageSetup.BlackAndWhite = true;
                }

                // Save the workbook to a file
                workbook.Save("Workbook_BlackAndWhite.xlsx");
                Console.WriteLine("Workbook saved successfully as Workbook_BlackAndWhite.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableBlackAndWhitePrinting.Run();
        }
    }
}
