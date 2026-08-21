// Title: Unhide Row 10 in the Active Worksheet and Save Workbook with Aspose.Cells for .NET
// Description: Load an existing Excel file, access the active sheet, unhide row 10 (zero‑based index 9) with auto‑fit, and save the modified workbook as a new file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# unhide row | unhide Excel row .NET | active worksheet | save workbook Aspose.Cells | auto fit row height | Excel row visibility | Aspose.Cells API
// Common Searches: Aspose.Cells unhide row 10 | C# hide/unhide rows Excel | How to unhide a row in Aspose.Cells | Save workbook after modifying rows Aspose.Cells | Active sheet index Aspose.Cells C#
// Developer Intent: Load an Excel workbook, make row 10 visible on the active sheet, and write the changes back to disk.
// Use Cases: Reveal a hidden header before exporting the sheet to PDF. | Prepare a template workbook by unhiding rows prior to data population. | Ensure all rows are displayed in a generated report for accurate printing.
// AI Prompts: Create C# code that opens an Excel file with Aspose.Cells, unhides rows 5‑10, sets a custom height for each row, and saves the result to a memory stream. | Show an example of using Aspose.Cells to unhide a specific row, auto‑fit its height, and then save the workbook in a different format such as CSV or PDF.

using System;
using Aspose.Cells;

// Load an existing Excel file, access the active sheet, unhide row 10 (zero‑based index 9) with auto‑fit, and save the modified workbook as a new file using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the active worksheet (the one currently selected)
        Worksheet worksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Unhide row 10 (zero‑based index 9) and let Aspose.Cells auto‑fit the height
        worksheet.Cells.UnhideRow(9, -1);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
