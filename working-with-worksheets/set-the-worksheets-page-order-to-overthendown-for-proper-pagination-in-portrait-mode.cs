// Title: C# – Set worksheet page order to OverThenDown (portrait) with Aspose.Cells
// Description: Shows how to create a workbook, switch the sheet to portrait mode, assign PrintOrderType.OverThenDown to PageSetup.Order, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# page order | PrintOrderType OverThenDown | worksheet pagination portrait | PageSetup.Order | Excel print order .NET | Aspose.Cells tutorial | set page order C#
// Common Searches: Aspose.Cells set page order OverThenDown | C# print order portrait Excel | How to change worksheet pagination order in Aspose.Cells | PrintOrderType OverThenDown example | Configure page setup orientation and order Aspose.Cells
// Developer Intent: Configure a worksheet's print order to OverThenDown while keeping the sheet in portrait mode in an Aspose.Cells workbook.
// Use Cases: Printing multi‑column reports where pages fill horizontally before moving down the sheet. | Generating invoices that require horizontal page sequencing in a vertical layout. | Preparing Excel sheets for booklet or brochure printing with OverThenDown pagination. | Automating batch export of reports where a consistent page order is mandatory.
// AI Prompts: Generate C# code with Aspose.Cells that sets PageSetup.Orientation to portrait, assigns PageSetup.Order = PrintOrderType.OverThenDown, and saves the workbook as XLSX. | Explain the difference between OverThenDown and DownThenOver print orders in Aspose.Cells and recommend scenarios for each. | Provide a step‑by‑step tutorial for configuring page setup (orientation, margins, scaling, order) in an Aspose.Cells workbook using .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, switch the sheet to portrait mode, assign PrintOrderType.OverThenDown to PageSetup.Order, and save the result as an XLSX file using Aspose.Cells for .NET.
    public class SetPageOrderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the sheet is in portrait orientation
                sheet.PageSetup.Orientation = PageOrientationType.Portrait;

                // Set the page order to OverThenDown for proper pagination
                sheet.PageSetup.Order = PrintOrderType.OverThenDown;

                // Define output file name
                string outputPath = "PageOrder_OverThenDown.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
