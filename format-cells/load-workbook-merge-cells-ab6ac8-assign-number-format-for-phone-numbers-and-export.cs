// Title: C# – Merge AB6:AC8 and Apply Phone Number Format (###) ###‑#### with Aspose.Cells
// Description: Load an existing XLSX file, merge the range AB6:AC8 on the first worksheet, assign the custom number format "(###) ###‑####" to the merged cell, and save the workbook as a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | custom phone number format Aspose | apply number format .NET | merge range AB6 AC8 | save workbook Aspose.Cells
// Common Searches: Aspose.Cells merge AB6 AC8 C# | set custom phone number format in Aspose.Cells | how to apply (###) ###‑#### format to merged cells | C# Aspose.Cells example save workbook after formatting
// Developer Intent: Merge cells AB6:AC8, set a phone‑number custom format, and save the workbook.
// Use Cases: Create a header spanning columns AB and AC that displays a formatted phone number. | Build a contact sheet where each merged cell holds a phone number in a consistent format. | Generate a report that requires merged cells to show phone numbers with a specific pattern.
// AI Prompts: Show a C# example that merges AB6:AC8 and applies the "(###) ###‑####" format with Aspose.Cells. | Provide code to load an XLSX, merge a range, set a custom phone number format, and save the file using Aspose.Cells for .NET. | Explain how to ensure a custom number format persists on a merged cell after saving with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an existing XLSX file, merge the range AB6:AC8 on the first worksheet, assign the custom number format "(###) ###‑####" to the merged cell, and save the workbook as a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells AB6:AC8 (zero‑based indices)
        // Row 6 -> index 5, Column AB -> index 27
        int firstRow = 5;       // Row 6
        int firstColumn = 27;   // Column AB
        int totalRows = 3;      // Rows 6,7,8
        int totalColumns = 2;   // Columns AB, AC
        worksheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

        // Apply phone number format to the merged cell (AB6)
        Cell mergedCell = worksheet.Cells[firstRow, firstColumn];
        Style style = mergedCell.GetStyle();
        style.Custom = "(###) ###-####";
        mergedCell.SetStyle(style);

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
