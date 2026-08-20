// Title: Aspose.Cells for .NET – Merge the first worksheet row (A‑Z) into a centered bold header
// Description: Creates a new workbook, merges cells A1:Z1 on the first worksheet, inserts "Report Header", applies center alignment, bold 14‑pt font, and saves the file as HeaderMerged.xlsx.
// Keywords: Aspose.Cells merge first row | C# merge cells A to Z | Excel header spanning columns | centered bold header Aspose | style merged cell .NET
// Common Searches: how to merge first row Aspose.Cells | Aspose.Cells create full‑width header | C# merge A1 Z1 and center text | Aspose.Cells set style for merged header | Excel report title across columns using Aspose
// Developer Intent: Generate a single merged title cell across the top row and format it as a centered bold header.
// Use Cases: Automated report generation with a full‑width title row | Dashboard worksheets that need a prominent header spanning all columns | Printable spreadsheets where the first row serves as a centered report heading
// AI Prompts: Show C# code with Aspose.Cells that merges A1:Z1, writes "Report Header", centers the text, makes it bold, and saves the workbook. | Provide an Aspose.Cells example for creating a merged header row and applying custom styling. | Explain how to modify the range to include additional rows or a different column span in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderMergeDemo
{
    // Creates a new workbook, merges cells A1:Z1 on the first worksheet, inserts "Report Header", applies center alignment, bold 14‑pt font, and saves the file as HeaderMerged.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the range to merge:
            // First row (index 0), starting at first column (index 0),
            // spanning 1 row and 26 columns (A to Z)
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = 1;      // one row
            int totalColumns = 26;  // columns A-Z

            // Merge the defined range to create a header
            cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

            // Set header text in the merged cell (upper‑left cell of the range)
            cells[firstRow, firstColumn].PutValue("Report Header");

            // Apply basic styling to the header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 14;
            cells[firstRow, firstColumn].SetStyle(headerStyle);

            // Save the workbook
            workbook.Save("HeaderMerged.xlsx");
        }
    }
}
