// Title: Apply a custom number format and save as tab‑delimited CSV with Aspose.Cells for .NET
// Description: Loads an Excel file, creates a style with a built‑in currency format, applies the style only to the number format of cells B2:B4, configures TxtSaveOptions to use a tab separator, and saves the workbook as a TSV file.
// Keywords: Aspose.Cells | C# | custom number format | currency style | apply style to range | tab delimited CSV | TSV export | TxtSaveOptions | SaveFormat.Csv | Excel to TSV
// Common Searches: Aspose.Cells set currency format for a range | C# export Excel to tab separated values | How to apply number format to cells with Aspose.Cells | Save workbook as TSV using Aspose.Cells | Tab delimiter CSV example Aspose.Cells
// Developer Intent: The developer needs C# code that formats a specific cell range with a custom number format and then exports the workbook as a tab‑separated CSV (TSV) using Aspose.Cells.
// Use Cases: Generate financial reports where monetary columns retain currency formatting in a TSV file. | Provide data to legacy systems that accept tab‑separated values while preserving numeric display. | Create downloadable TSV files from Excel templates after applying targeted formatting. | Automate data pipelines that require both custom formatting and a tab‑delimited output.
// AI Prompts: Write C# code with Aspose.Cells to apply a currency number format to cells B2:B4 and save the workbook as a tab‑delimited CSV. | Explain how TxtSaveOptions can be configured for tab separation while keeping number formats intact in Aspose.Cells. | Show an example of setting a date format for a range and exporting the workbook as a TSV file using Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Loads an Excel file, creates a style with a built‑in currency format, applies the style only to the number format of cells B2:B4, configures TxtSaveOptions to use a tab separator, and saves the workbook as a TSV file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create a style with a custom number format (e.g., currency)
        Style customStyle = workbook.CreateStyle();
        customStyle.Number = 5; // Currency format

        // Apply the style to a specific range (e.g., B2:B4)
        Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange("B2:B4");
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply only the number format
        range.ApplyStyle(customStyle, flag);

        // Configure text save options for tab‑delimited CSV
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.Separator = '\t';

        // Save the workbook as a TSV (tab‑delimited) file
        workbook.Save("output.tsv", txtOptions);
    }
}
