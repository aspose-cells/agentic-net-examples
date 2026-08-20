// Title: C# – Merge Header Cells (A1:D1), Set 14‑Point Font, and Save Workbook with Aspose.Cells
// Description: Load an existing Excel file, merge cells A1 through D1 on the first worksheet to create a single header, apply a 14‑point font to the merged cell, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | header row formatting Aspose.Cells | set font size 14 Aspose.Cells | Excel workbook export .NET | format cells Aspose.Cells | Aspose.Cells .NET US | Aspose.Cells Europe | C# Excel styling
// Common Searches: how to merge cells A1:D1 with Aspose.Cells C# | set font size of merged header cell Aspose.Cells .NET | save workbook after formatting header row Aspose.Cells | Aspose.Cells merge cells and change font size example | C# code to create a title header in Excel using Aspose
// Developer Intent: Combine the first‑row cells into one header, apply a 14‑point font, and write the changes back to a new Excel file.
// Use Cases: Generate report sheets with a bold title spanning multiple columns. | Standardize header appearance across automated Excel exports. | Create printable dashboards where the header needs larger, readable text.
// AI Prompts: Show C# code that merges A1:D1 and sets the font size to 14 with Aspose.Cells. | Give an Aspose.Cells example for styling a merged header row and saving the workbook. | Explain how to change font attributes (size, bold, color) of a merged cell after merging using Aspose.Cells in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderMerge
{
    // Load an existing Excel file, merge cells A1 through D1 on the first worksheet to create a single header, apply a 14‑point font to the merged cell, and save the updated workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells in the header row (row 0, columns A to D)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(0, 0, 1, 4);

            // Set the font size of the merged header cell to 14
            Style headerStyle = cells[0, 0].GetStyle();
            headerStyle.Font.Size = 14;
            cells[0, 0].SetStyle(headerStyle);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}
