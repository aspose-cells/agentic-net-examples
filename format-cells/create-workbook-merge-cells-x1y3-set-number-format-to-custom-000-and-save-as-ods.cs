// Title: Aspose.Cells for .NET: Merge X1:Y3, apply Euro custom format, and save as ODS
// Description: Creates a new workbook, merges cells X1:Y3, sets the custom number format "#,##0.00 €" on the merged cell, configures OdsSaveOptions for the LibreOffice generator, and saves the file as MergedFormatted.ods.
// Keywords: Aspose.Cells | C# | .NET | merge cells | X1:Y3 | custom number format | Euro currency | ODS export | OdsSaveOptions | LibreOffice generator | merged cell styling
// Common Searches: How to merge cells X1 to Y3 in Aspose.Cells C# | Set Euro currency format "#,##0.00 €" on merged cells with Aspose.Cells | Save Aspose.Cells workbook as ODS using LibreOffice generator | Apply custom number format to a merged range in .NET spreadsheet | Aspose.Cells example for merged header with Euro format in ODS
// Developer Intent: Create an ODS spreadsheet with a merged header (X1:Y3) formatted in Euro currency.
// Use Cases: Generating financial report headers that span multiple columns and display amounts in Euro before sharing with LibreOffice users. | Automating invoice templates where the title cell is merged and styled with a Euro currency pattern for cross‑platform compatibility. | Building spreadsheet export features that require merged cells with custom currency formatting for multinational accounting systems. | Preparing dashboard labels where a merged cell serves as a localized Euro‑denominated heading.
// AI Prompts: Provide C# code using Aspose.Cells to merge the range X1:Y3, apply the custom number format "#,##0.00 €" to the merged cell, and save the workbook as an ODS file with the LibreOffice generator. | Show how to style the upper‑left cell of a merged range with a Euro currency format and persist the style when exporting to ODS using Aspose.Cells for .NET. | Explain the steps to configure OdsSaveOptions for LibreOffice compatibility while merging cells and setting a custom currency format in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new workbook, merges cells X1:Y3, sets the custom number format "#,##0.00 €" on the merged cell, configures OdsSaveOptions for the LibreOffice generator, and saves the file as MergedFormatted.ods.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells X1:Y3 (zero‑based indices: row 0, column 23 for X; span 3 rows, 2 columns)
        worksheet.Cells.Merge(0, 23, 3, 2);

        // Apply a custom number format to the merged cell (upper‑left cell of the range)
        Cell mergedCell = worksheet.Cells["X1"];
        // Use a custom format that includes the Euro symbol
        mergedCell.GetStyle().Custom = "#,##0.00 €";
        mergedCell.SetStyle(mergedCell.GetStyle());

        // Prepare ODS save options
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };

        // Save the workbook as ODS
        workbook.Save("MergedFormatted.ods", saveOptions);
    }
}
