// Title: Apply a uniform StandardHeight to every worksheet in an Aspose.Cells workbook (C#)
// Description: Creates a new Workbook, optionally adds extra sheets, defines a single row height value, loops through all worksheets and sets Cells.StandardHeight to enforce the same default row height across the entire workbook, then saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | StandardHeight | row height | uniform row height | default row height | iterate worksheets | Workbook | Cells.StandardHeight property | Excel row height | Aspose.Cells API
// Common Searches: Aspose.Cells set same row height for all sheets | C# iterate worksheets set StandardHeight | default row height Aspose.Cells workbook | how to apply uniform row height in Aspose.Cells | set Cells.StandardHeight across multiple worksheets
// Developer Intent: Set a consistent default row height for every worksheet in a workbook using Aspose.Cells for .NET.
// Use Cases: Create a reporting template where every sheet must share a 20‑point row height before data is added. | Standardize row height in an existing workbook to ensure uniform appearance when exporting to PDF or printing. | Generate a multi‑sheet Excel file programmatically and enforce identical row height across all sheets for visual consistency.
// AI Prompts: Write C# code with Aspose.Cells that opens an existing workbook, sets Cells.StandardHeight to 15 points for all worksheets, and saves the result as .xlsx. | Provide a reusable method that accepts a Workbook object and a double height, iterates through each Worksheet, applies the StandardHeight, and returns the updated workbook. | Explain the difference between using Cells.StandardHeight and setting individual Row.Height values, and advise when each approach is preferable.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, optionally adds extra sheets, defines a single row height value, loops through all worksheets and sets Cells.StandardHeight to enforce the same default row height across the entire workbook, then saves the file as an XLSX document.
    public class SetStandardRowHeightForAllSheets
    {
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
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Optionally add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Desired uniform row height in points
            double uniformHeight = 20.0;

            // Iterate through each worksheet and set the StandardHeight (property rule)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // StandardHeight sets the default row height for the entire worksheet
                sheet.Cells.StandardHeight = uniformHeight;
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("UniformRowHeightWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
