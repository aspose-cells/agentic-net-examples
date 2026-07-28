// Title: Merge C6:E7, Apply Bold Font, and Save as XLSX with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX file, merges the range C6:E7 on the first worksheet, sets the top‑left cell of the merged area to bold, and saves the modified workbook as a new XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# merge cells | merge C6:E7 | bold font style | save workbook XLSX | Aspose.Cells .NET | cell merging Aspose | format merged cells | Excel automation C# | SaveFormat.Xlsx
// Common Searches: Aspose.Cells merge cells C6:E7 C# | How to make merged cells bold with Aspose.Cells | Save workbook after merging cells Aspose.Cells .NET | C# Aspose.Cells merge range and apply style | Aspose.Cells example merge and bold
// Developer Intent: Combine a specific cell range, apply bold formatting, and export the workbook as an XLSX file.
// Use Cases: Create a centered header spanning columns C‑E in a financial report and highlight it with bold text. | Generate an invoice template where the title row is merged across C‑E and displayed in bold. | Prepare a data‑entry template that requires merged header cells with bold styling before distribution.
// AI Prompts: Generate C# code using Aspose.Cells to merge cells C6:E7, set the font to bold, and save the workbook as output.xlsx. | Explain how to apply a style to the top‑left cell after merging a range with Aspose.Cells for .NET. | Provide best‑practice error handling for loading, merging, styling, and saving an Excel workbook with Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace MergeCellsExample
{
    // Loads an existing XLSX file, merges the range C6:E7 on the first worksheet, sets the top‑left cell of the merged area to bold, and saves the modified workbook as a new XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook file
            string inputPath = "input.xlsx";

            // Load the workbook from disk using the provided constructor
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells C6:E7
            // C -> column index 2, row 6 -> row index 5 (zero‑based)
            // Total rows = 2 (rows 6 and 7), total columns = 3 (C, D, E)
            cells.Merge(5, 2, 2, 3);

            // Apply bold font to the merged cell (top‑left cell of the range)
            Style style = cells[5, 2].GetStyle();
            style.Font.IsBold = true;
            cells[5, 2].SetStyle(style);

            // Save the modified workbook as XLSX using the provided Save method
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
