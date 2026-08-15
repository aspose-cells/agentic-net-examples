// Title: Aspose.Cells for .NET: Create a hidden worksheet and invisible named range for internal calculations
// Description: Demonstrates how to add a hidden worksheet, populate cells, define an invisible named range that points to that sheet, and use the name in a SUM formula on a visible sheet before saving the workbook.
// Keywords: Aspose.Cells | .NET | C# | hidden worksheet | named range | invisible name | internal calculations | Excel formula | SUM function | Workbook automation
// Common Searches: Aspose.Cells create hidden sheet C# | define invisible named range Aspose.Cells | reference hidden worksheet range in formula | hide named range from Excel UI | internal calculation sheet Aspose.Cells
// Developer Intent: Add a concealed worksheet and a non‑visible named range that references it, then use that name in formulas on other sheets.
// Use Cases: Store lookup tables on a hidden sheet and call them via an invisible name in visible‑sheet formulas. | Keep intermediate results out of the user view while still participating in calculations. | Create reusable internal data ranges that are hidden from end‑users but available to the workbook logic.
// AI Prompts: Generate C# code with Aspose.Cells that adds a hidden worksheet, fills A1:A5, creates an invisible named range referencing that range, and uses it in a SUM formula on another sheet. | Explain how setting Name.IsVisible = false affects workbook users and how to access the hidden range programmatically. | Show an example of using a hidden named range for internal calculations and then saving the workbook as an .xlsx file.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a hidden worksheet, populate cells, define an invisible named range that points to that sheet, and use the name in a SUM formula on a visible sheet before saving the workbook.
    public class HiddenWorksheetNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a hidden worksheet for internal calculations
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenCalc");
                // Hide the worksheet
                hiddenSheet.IsVisible = false;

                // Populate some data on the hidden sheet (e.g., A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    hiddenSheet.Cells[i, 0].PutValue(i + 1); // Values 1..5 in column A
                }

                // Create a named range that refers to the range on the hidden worksheet
                int nameIndex = workbook.Worksheets.Names.Add("HiddenValues");
                Name hiddenRangeName = workbook.Worksheets.Names[nameIndex];
                // Set the reference to the hidden sheet range (absolute reference)
                hiddenRangeName.RefersTo = "=HiddenCalc!$A$1:$A$5";

                // Optionally make the name invisible to users
                hiddenRangeName.IsVisible = false;

                // Example usage: use the named range in a formula on the first (visible) sheet
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "Sheet1";
                visibleSheet.Cells["B1"].Formula = "=SUM(HiddenValues)";
                workbook.CalculateFormula();

                // Output the result of the formula to console
                Console.WriteLine("Sum of hidden values: " + visibleSheet.Cells["B1"].Value);

                // Save the workbook
                workbook.Save("HiddenWorksheetNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HiddenWorksheetNamedRangeDemo.Run();
        }
    }
}
