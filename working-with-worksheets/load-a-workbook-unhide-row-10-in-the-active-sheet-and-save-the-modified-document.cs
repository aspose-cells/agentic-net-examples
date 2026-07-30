// Title: C# – Unhide Row 10 in the Active Worksheet using Aspose.Cells and Save the Workbook
// Description: Loads an existing Excel file, accesses the active sheet, unhides row 10 (zero‑based index 9) with automatic height adjustment, and saves the result to a new file.
// Keywords: Aspose.Cells unhide row C# | unhide specific row Aspose.Cells | active worksheet row visibility | save workbook after row change | auto‑fit row height Aspose.Cells
// Common Searches: Aspose.Cells C# unhide row 10 | how to make a hidden row visible with Aspose.Cells | retrieve active worksheet and modify row visibility | save Excel file after changing row visibility Aspose
// Developer Intent: Programmatically reveal row 10 in the currently active worksheet and write the updated workbook to disk.
// Use Cases: Expose a hidden header row before exporting a report. | Make a specific data row visible for downstream processing after loading a template. | Adjust row visibility in a dynamic Excel template prior to populating it with content.
// AI Prompts: Generate C# code with Aspose.Cells to unhide rows 5‑15 and save the workbook. | Show how to unhide a row and set a custom height using Aspose.Cells for .NET. | Provide an example that checks if a row is hidden before calling UnhideRow in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsRowUnhideExample
{
    // Loads an existing Excel file, accesses the active sheet, unhides row 10 (zero‑based index 9) with automatic height adjustment, and saves the result to a new file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Get the active worksheet (the one currently selected)
            Worksheet worksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Unhide row 10 (zero‑based index 9) and let Aspose.Cells auto‑fit the height
            worksheet.Cells.UnhideRow(9, -1);

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
