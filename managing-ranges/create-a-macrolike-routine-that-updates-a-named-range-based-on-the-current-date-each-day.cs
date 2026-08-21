// Title: Update a Named Range with Today’s Date using Aspose.Cells for .NET
// Description: Demonstrates a macro‑like routine that creates a workbook, defines a named range (e.g., "TodayDate"), fills the range with the current date, applies a standard date format, and saves the file as UpdatedNamedRange.xlsx.
// Keywords: Aspose.Cells C# named range date | set named range to today Aspose.Cells | update named range programmatically .NET | apply date format Aspose.Cells | macro‑like routine Aspose.Cells
// Common Searches: Aspose.Cells set named range to current date | C# update named range daily Aspose.Cells | how to apply date format to named range in Aspose.Cells | macro alternative for updating date cell Aspose.Cells | schedule Aspose.Cells code to run each day
// Developer Intent: Provide sample code that programmatically refreshes a named range with the current date and proper formatting in a .NET workbook.
// Use Cases: Automated daily reports where the header shows the generation date. | Workbook templates that display today’s date whenever opened or saved. | Dashboard files that need the date cell refreshed before publishing.
// AI Prompts: Generate C# Aspose.Cells code to create a named range "ReportDate" at cell B2, insert DateTime.Now, apply a custom date format, and save the workbook. | Show how to loop through a multi‑cell named range and set each cell to the current date with a specific style using Aspose.Cells. | Explain how to schedule the UpdateNamedRangeWithCurrentDate method to execute automatically each day on Windows or Linux.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates a macro‑like routine that creates a workbook, defines a named range (e.g., "TodayDate"), fills the range with the current date, applies a standard date format, and saves the file as UpdatedNamedRange.xlsx.
    public class UpdateNamedRangeWithCurrentDate
    {
        // Entry point for the example
        public static void Main()
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

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell that will hold the date (e.g., A1)
            string targetCellAddress = "A1";

            // Create a named range called "TodayDate" that refers to the target cell
            int nameIndex = workbook.Worksheets.Names.Add("TodayDate");
            Name todayName = workbook.Worksheets.Names[nameIndex];
            todayName.RefersTo = $"={sheet.Name}!${targetCellAddress}";

            // Retrieve the range that the name points to
            Aspose.Cells.Range dateRange = todayName.GetRange();

            // Create a style for date formatting (Number format 14 = "m/d/yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14;

            // Update each cell in the range with the current date and apply the style
            foreach (Cell cell in dateRange)
            {
                cell.PutValue(DateTime.Today);
                cell.SetStyle(dateStyle);
            }

            // Save the workbook to a file
            string outputPath = "UpdatedNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
