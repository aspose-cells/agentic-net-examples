// Title: Load an Excel file, merge cells C6:E7, apply bold formatting, and save as XLSX with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells merge a range of cells and set bold font then save as XLSX | How to programmatically merge cells C6 to E7 in an existing workbook using Aspose.Cells .NET | Apply bold style to a merged cell area with Aspose.Cells in C# | Load workbook from disk, merge cells, and export to XLSX using Aspose.Cells for .NET | Aspose.Cells .NET example for merging cells and applying font style
// Tags: merge cell range C6:E7 Aspose.Cells | bold font style for merged cells Aspose.Cells | load workbook from disk Aspose.Cells .NET | save workbook as XLSX Aspose.Cells | CellArea merge operation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an existing Excel workbook (or creating a new one), merging cells C6 through E7, applying a bold font style to the merged range, and saving the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                // Load existing workbook (no password required)
                LoadOptions loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // Create a new workbook if the input file is missing
                workbook = new Workbook();
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge: C6:E7 (zero‑based indexes)
            CellArea mergeArea = new CellArea
            {
                StartRow = 5,   // Row 6 in Excel
                EndRow = 6,     // Row 7 in Excel
                StartColumn = 2, // Column C in Excel
                EndColumn = 4   // Column E in Excel
            };

            // Merge the cells
            sheet.Cells.Merge(
                mergeArea.StartRow,
                mergeArea.StartColumn,
                mergeArea.EndRow - mergeArea.StartRow + 1,
                mergeArea.EndColumn - mergeArea.StartColumn + 1);

            // Create a style with bold font
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            StyleFlag flag = new StyleFlag
            {
                FontBold = true // Apply only the bold attribute
            };

            // Apply the style to the merged range
            sheet.Cells.CreateRange(
                mergeArea.StartRow,
                mergeArea.StartColumn,
                mergeArea.EndRow - mergeArea.StartRow + 1,
                mergeArea.EndColumn - mergeArea.StartColumn + 1)
                .ApplyStyle(style, flag);

            // Save the workbook as XLSX
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
