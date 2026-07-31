// Title: Merge X1:Y3, apply Euro currency format, and save as ODS with Aspose.Cells for .NET
// Description: Creates a new Workbook, merges the range X1:Y3, defines a custom number format "#,##0.00 €", applies the style to the merged area, and saves the file as an ODS document using OdsSaveOptions.
// Keywords: Aspose.Cells | C# | .NET | merge cells | X1:Y3 | custom number format | Euro currency | ODS export | OdsSaveOptions | financial reporting
// Common Searches: Aspose.Cells merge cells X1 Y3 C# | custom Euro number format Aspose.Cells | save workbook as ODS .NET | apply style to merged range Aspose.Cells | ODS output with currency formatting
// Developer Intent: Generate an ODS file where cells X1:Y3 are merged and displayed with a Euro currency format.
// Use Cases: Financial statements with merged total rows formatted in euros. | Invoice templates that require a pre‑merged header cell showing currency values. | Exporting calculation results to ODS while preserving custom monetary formatting.
// AI Prompts: Show how to merge cells X1:Y3 and set a "#,##0.00 €" format using Aspose.Cells for .NET. | Provide a C# example that saves a workbook with merged, Euro‑formatted cells as an ODS file. | Explain the steps to apply a custom currency number format to a merged range and export it to ODS with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new Workbook, merges the range X1:Y3, defines a custom number format "#,##0.00 €", applies the style to the merged area, and saves the file as an ODS document using OdsSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells X1:Y3 (zero‑based indices: row 0, column 23, 3 rows, 2 columns)
        worksheet.Cells.Merge(0, 23, 3, 2);

        // Create a style with a custom number format '#,##0.00 €'
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0.00 €";

        // Apply the style to the merged cell (upper‑left cell of the range)
        worksheet.Cells[0, 23].SetStyle(customStyle);

        // Prepare ODS save options (default options are sufficient)
        OdsSaveOptions saveOptions = new OdsSaveOptions();

        // Save the workbook as ODS
        workbook.Save("MergedFormatted.ods", saveOptions);
    }
}
