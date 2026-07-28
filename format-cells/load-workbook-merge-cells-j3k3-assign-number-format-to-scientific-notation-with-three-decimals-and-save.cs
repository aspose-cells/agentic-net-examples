// Title: C# – Merge J3:K3 and apply 3‑decimal scientific format with Aspose.Cells
// Description: Shows how to merge cells J3:K3 in a worksheet, assign the custom number format 0.000E+00 to display values in scientific notation with three decimal places, add a sample value, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | C# scientific notation format | custom number format 0.000E+00 | merged cell formatting .NET | Excel J3 K3 merge Aspose | three decimal scientific format
// Common Searches: Aspose.Cells merge cells J3 K3 C# | set scientific notation format Aspose.Cells | custom number format for merged range .NET | C# code 0.000E+00 format Aspose.Cells | preserve cell format after saving Aspose workbook
// Developer Intent: Programmatically merge a specific range and apply a three‑decimal scientific number format.
// Use Cases: Create a report header that spans two columns and automatically shows values in scientific notation. | Prepare exported data where merged cells must retain a consistent scientific format across different viewers. | Build a spreadsheet template with pre‑merged cells that enforce a 3‑decimal scientific display for user‑entered numbers.
// AI Prompts: Provide C# code to merge cells J3:K3 and set the custom format 0.000E+00 using Aspose.Cells. | How can I apply a three‑decimal scientific notation to the upper‑left cell of a merged range in Aspose.Cells for .NET? | Explain how to ensure a custom number format persists after saving a workbook that contains merged cells with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to merge cells J3:K3 in a worksheet, assign the custom number format 0.000E+00 to display values in scientific notation with three decimal places, add a sample value, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells J3:K3 (zero‑based indices: row 2, column 9, 1 row, 2 columns)
        worksheet.Cells.Merge(2, 9, 1, 2);

        // Apply scientific notation with three decimal places to the merged cell
        Cell mergedCell = worksheet.Cells[2, 9]; // Upper‑left cell of the merged range
        Style style = mergedCell.GetStyle();
        // Use a custom format string for three‑decimal scientific notation
        style.Custom = "0.000E+00";
        mergedCell.SetStyle(style);

        // Optionally put a sample value to see the formatting
        mergedCell.PutValue(12345.6789);

        // Save the workbook
        workbook.Save("MergedScientific.xlsx");
    }
}
