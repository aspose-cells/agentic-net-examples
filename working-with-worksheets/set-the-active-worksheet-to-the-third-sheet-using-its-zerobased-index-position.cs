// Title: C# Aspose.Cells: Set Active Worksheet to Third Sheet (ActiveSheetIndex = 2)
// Description: Creates a Workbook, adds two extra worksheets, assigns Worksheets.ActiveSheetIndex = 2 to make the third sheet active, outputs its name, and saves the workbook as ActiveSheetSet.xlsx.
// Keywords: Aspose.Cells | ActiveSheetIndex | C# | .NET | set active worksheet | third worksheet | zero‑based index | Workbook.Worksheets | Excel automation
// Common Searches: Aspose.Cells set active sheet C# | ActiveSheetIndex example .NET | Select third worksheet in Aspose.Cells | Change active worksheet by index C# | Get active sheet name Aspose.Cells
// Developer Intent: Activate the third worksheet in a workbook by using its zero‑based index.
// Use Cases: Display a specific sheet when the file opens in Excel. | Apply formatting or formulas to a target sheet before saving. | Export data from a chosen worksheet to another format. | Programmatically switch sheets during batch processing.
// AI Prompts: Show a C# snippet that sets the active worksheet to the third sheet using Aspose.Cells and confirms the change. | Provide code that changes the active sheet based on a variable index and saves the workbook. | Explain how to retrieve the name of the currently active worksheet after modifying ActiveSheetIndex.

using System;
using Aspose.Cells;

namespace AsposeCellsActiveSheetDemo
{
    // Creates a Workbook, adds two extra worksheets, assigns Worksheets.ActiveSheetIndex = 2 to make the third sheet active, outputs its name, and saves the workbook as ActiveSheetSet.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add two more worksheets so we have at least three sheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Set the active worksheet to the third sheet (zero‑based index 2)
            workbook.Worksheets.ActiveSheetIndex = 2;

            // Optional: verify the active sheet name
            Console.WriteLine("Active Sheet: " + workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex].Name);

            // Save the workbook to a file
            workbook.Save("ActiveSheetSet.xlsx");
        }
    }
}
