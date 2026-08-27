// Title: Apply a custom number format to a range and save the workbook as a tab‑delimited CSV using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing Excel file, creates a custom number style (e.g., "#,##0.00"), applies it to a specific cell range, and saves the workbook as a TSV file with tab delimiters using Aspose.Cells. | Show how to configure TxtSaveOptions to preserve display formatting when exporting a worksheet to a tab‑delimited CSV in Aspose.Cells for .NET.
// Common Searches: asp.net how to export Excel to tab delimited CSV with custom number formatting using Aspose.Cells | c# set custom number format for cells before saving as TSV with Aspose.Cells | preserve cell display style when converting workbook to tab separated values in Aspose.Cells | txtsaveoptions separator tab example Aspose.Cells .NET | apply number style to range B2:B5 Aspose.Cells C#
// Tags: custom number format Aspose.Cells | apply style to cell range C# | tab delimited CSV export Aspose.Cells | tab delimiter TxtSaveOptions | preserve display formatting Aspose.Cells | SaveFormat.Csv Aspose.Cells example

using System;
using Aspose.Cells;

// The sample loads 'input.xlsx', creates a custom number style '#,##0.00', applies it to cells B2:B5, fills those cells with numeric values, configures TxtSaveOptions with a tab separator and DisplayStyle strategy to retain formatting, and saves the result as 'output.tsv' (a tab‑delimited CSV).
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a custom number style (e.g., two decimal places with thousand separator)
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0.00";

        // Apply the style to a range of cells (B2:B5 in this example)
        worksheet.Cells.CreateRange("B2:B5").SetStyle(customStyle);

        // Populate the range with sample numeric values
        worksheet.Cells["B2"].PutValue(1234.56);
        worksheet.Cells["B3"].PutValue(7890);
        worksheet.Cells["B4"].PutValue(0.123);
        worksheet.Cells["B5"].PutValue(45678.9);

        // Configure text save options for tab‑delimited output
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = '\t';                     // Tab delimiter
        saveOptions.FormatStrategy = CellValueFormatStrategy.DisplayStyle; // Preserve formatting

        // Save the workbook as a TSV (tab‑delimited) file
        workbook.Save("output.tsv", saveOptions);
    }
}
