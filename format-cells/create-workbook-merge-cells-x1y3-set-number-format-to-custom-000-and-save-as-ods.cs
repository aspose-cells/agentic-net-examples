// Title: How to merge cells X1:Y3, apply a custom Euro currency format, and save as ODS using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, merges the range X1:Y3, applies the custom number format '#,##0.00 €' to the merged cells, and saves the file in OpenDocument Spreadsheet (ODS) format with Aspose.Cells. | Write a C# example that uses Aspose.Cells to merge a specific cell block, set a Euro currency custom format on that range, and export the workbook as an ODS file.
// Common Searches: Aspose.Cells C# merge specific cells and set custom currency format | Save merged cells with Euro format to ODS using Aspose.Cells | C# example for applying '#,##0.00 €' number format to a merged range in Aspose.Cells | How to export a workbook with formatted merged cells to OpenDocument Spreadsheet in .NET
// Tags: merge cells range Aspose.Cells C# | custom number format Euro Aspose.Cells | export workbook to ODS Aspose.Cells | apply style to merged range Aspose.Cells | open document spreadsheet format Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, merges cells X1:Y3, applies the custom Euro currency number format '#,##0.00 €' to the merged range, and saves the file as an ODS document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells X1:Y3 (zero‑based indices: rows 0‑2, columns 23‑24)
            sheet.Cells.Merge(0, 23, 3, 2);

            // Create a style with the custom number format '#,##0.00 €'
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0.00 €";

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange(0, 23, 3, 2);
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            mergedRange.ApplyStyle(customStyle, flag);

            // Save the workbook as ODS
            workbook.Save("MergedCells.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
