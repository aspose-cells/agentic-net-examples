// Title: Hide First Worksheet Tab & Show Formulas on Second Sheet with Aspose.Cells for .NET
// Description: Creates a workbook, hides the first worksheet tab, ensures a second sheet exists, enables ShowFormulas on that sheet, and saves the file as an .xlsx document using Aspose.Cells for C#.
// Keywords: Aspose.Cells hide worksheet | ShowFormulas C# | Aspose.Cells hide first sheet | display formulas Aspose.Cells | save workbook .xlsx | C# workbook visibility
// Common Searches: Aspose.Cells hide first worksheet tab C# | Show formulas on a specific sheet with Aspose.Cells | How to save a workbook after changing sheet visibility in .NET | Enable formula view for second worksheet using Aspose.Cells
// Developer Intent: Hide the first sheet, display formulas on the second sheet, and persist the workbook.
// Use Cases: Prepare a template where a configuration sheet is hidden while the calculation sheet shows formulas for audit trails. | Distribute a report that conceals internal data on the first tab but reveals the underlying formulas on the next tab for reviewers. | Create a workbook for training that protects the introductory sheet and lets learners see formula logic on the subsequent sheet.
// AI Prompts: Generate C# code with Aspose.Cells that hides the first worksheet, adds a second worksheet if missing, sets ShowFormulas = true on it, and saves the workbook as an .xlsx file. | Provide an Aspose.Cells example that checks worksheet count, hides the first tab, enables formula display on the second sheet, and writes the changes to disk.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, hides the first worksheet tab, ensures a second sheet exists, enables ShowFormulas on that sheet, and saves the file as an .xlsx document using Aspose.Cells for C#.
    public class HideFirstShowFormulasSecond
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook with default two worksheets
            Workbook workbook = new Workbook();

            // Hide the first worksheet tab
            workbook.Worksheets[0].IsVisible = false;

            // Ensure there is a second worksheet; add if missing
            if (workbook.Worksheets.Count < 2)
            {
                workbook.Worksheets.Add("Sheet2");
            }

            // Display formulas instead of calculated results on the second worksheet
            workbook.Worksheets[1].ShowFormulas = true;

            // Save the workbook with the applied changes
            workbook.Save("HiddenFirst_ShowFormulasSecond.xlsx", SaveFormat.Xlsx);
        }
    }
}
