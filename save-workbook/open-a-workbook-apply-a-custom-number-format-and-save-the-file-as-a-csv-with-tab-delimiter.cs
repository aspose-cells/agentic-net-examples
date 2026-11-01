// Title: How to apply a custom number format to a range of cells and export the workbook as a tab‑delimited CSV using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, creates a style with a custom number format (e.g., "#,##0.00;[Red]-#,##0.00;\"Zero\""), applies the style to cells A1:A10, and saves the workbook as a TSV file using TxtSaveOptions. | Show the steps to configure TxtSaveOptions for CSV output with a tab separator and apply a custom number format to a worksheet range before saving with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# apply custom numeric format to a column and save as TSV | How to set a custom number style for a range and export to tab‑separated CSV with Aspose.Cells | C# Aspose.Cells TxtSaveOptions tab delimiter example | Export Excel to tab delimited CSV while preserving cell formatting using Aspose.Cells
// Tags: custom numeric style Aspose.Cells | apply style to range Aspose.Cells | TxtSaveOptions tab separator | export to TSV Aspose.Cells | C# Aspose.Cells CSV with tab delimiter

using Aspose.Cells;
using System;

// Loads 'input.xlsx', defines a custom numeric style (#,##0.00;[Red]-#,##0.00;"Zero"), applies it to A1:A10, sets TxtSaveOptions with a tab separator, and saves the result as 'output.csv'.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet sheet = workbook.Worksheets[0];

        // Create a custom number format style
        Style customStyle = workbook.CreateStyle();
        // Example custom format: positive numbers with two decimals, negatives in red, zero displayed as "Zero"
        customStyle.Custom = "#,##0.00;[Red]-#,##0.00;\"Zero\"";

        // Apply the custom style to a range of cells (e.g., A1:A10)
        for (int row = 0; row < 10; row++)
        {
            Cell cell = sheet.Cells[row, 0]; // Column A
            cell.SetStyle(customStyle);
        }

        // Configure CSV save options with a tab delimiter
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        csvOptions.Separator = '\t'; // Tab delimiter

        // Save the workbook as a CSV file
        workbook.Save("output.csv", csvOptions);
    }
}
