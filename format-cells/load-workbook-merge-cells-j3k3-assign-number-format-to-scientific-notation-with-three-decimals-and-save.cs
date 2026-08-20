// Title: C# – Merge J3:K3 and Apply 3‑Decimal Scientific Notation in Aspose.Cells .NET
// Description: Loads an existing workbook, merges the range J3:K3 on the first worksheet, creates a style with the custom format "0.000E+00" (three decimal places in scientific notation), applies the style to the merged cell, writes a sample value, and saves the file as output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | merge cells | J3:K3 | scientific notation | custom number format | 0.000E+00 | cell style | Excel automation .NET | format merged cells
// Common Searches: Aspose.Cells merge cells C# | set scientific notation format Aspose.Cells | custom number format 0.000E+00 .NET | apply style to merged range Aspose.Cells | C# code merge J3 K3 Excel
// Developer Intent: Merge the range J3:K3, apply a three‑decimal scientific notation (0.000E+00) style, and save the workbook with Aspose.Cells for .NET.
// Use Cases: Create a header that spans columns J and K and displays values in scientific notation for engineering reports. | Prepare a financial model where specific merged cells must show high‑precision exponential values. | Generate an export file that requires both cell merging and a consistent custom number format across worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that merges J3:K3 and sets the number format to "0.000E+00". | Show how to apply a three‑decimal scientific notation style to a merged cell range in Aspose.Cells for .NET. | Explain how to reuse a Style object to format multiple merged cells with the same scientific notation in an Excel workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing workbook, merges the range J3:K3 on the first worksheet, creates a style with the custom format "0.000E+00" (three decimal places in scientific notation), applies the style to the merged cell, writes a sample value, and saves the file as output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path if needed)
            // If you want to create a new workbook, use: new Workbook();
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells J3:K3
            // J -> column index 9 (zero‑based), row 3 -> row index 2
            // Merge 1 row and 2 columns
            worksheet.Cells.Merge(2, 9, 1, 2);

            // Apply scientific notation with three decimal places to the merged cell
            Cell mergedCell = worksheet.Cells["J3"];
            Style style = mergedCell.GetStyle();

            // Use a custom number format for three decimal places in scientific notation
            style.Custom = "0.000E+00";

            mergedCell.SetStyle(style);

            // Optionally put a numeric value to demonstrate the format
            mergedCell.PutValue(12345.6789);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
    }
}
