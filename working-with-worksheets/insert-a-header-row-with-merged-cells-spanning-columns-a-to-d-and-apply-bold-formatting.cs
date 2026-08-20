// Title: C# – Insert a Bold, Centered Merged Header (A‑D) Using Aspose.Cells
// Description: Creates a new workbook, places a title in A1, merges cells A1:D1, applies bold font and center alignment, and saves the file as HeaderMerged.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | merged header row Aspose.Cells | bold centered text Aspose.Cells | C# Excel header formatting | Aspose.Cells workbook example
// Common Searches: how to merge A1 to D1 in Aspose.Cells | set bold centered header in Excel with C# | Aspose.Cells create merged title row | C# Aspose.Cells style merged cells
// Developer Intent: Generate a single‑row header that spans columns A‑D, merges the cells, and formats the text as bold and centered.
// Use Cases: Standard report templates where the first row shows a centered title across the first four columns. | Invoice or receipt generation with the company name displayed as a bold, merged header. | Dashboard worksheets that need a prominent, styled heading before data tables.
// AI Prompts: Write C# code with Aspose.Cells to merge A1:D1, set a title, apply bold font, center the text, and save the workbook. | Explain how to add a background color and increase the font size of a merged header row in Aspose.Cells for .NET. | Provide step‑by‑step instructions to create a merged header row with custom styling (font size, color, alignment) using Aspose.Cells.

using Aspose.Cells;

// Creates a new workbook, places a title in A1, merges cells A1:D1, applies bold font and center alignment, and saves the file as HeaderMerged.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put header text in the first cell (A1)
        cells[0, 0].PutValue("Header Title");

        // Merge cells A1:D1 (row 0, columns 0‑3)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        cells.Merge(0, 0, 1, 4);

        // Retrieve the style of the merged cell and apply bold formatting
        Style style = cells[0, 0].GetStyle();
        style.Font.IsBold = true;
        // Optional: center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;
        cells[0, 0].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("HeaderMerged.xlsx");
    }
}
