// Title: C# – Merge A2:D2 into a Centered Bold Header with Aspose.Cells for .NET
// Description: Creates a new workbook, merges cells A2:D2 on the first worksheet, sets the value "Header", applies horizontal and vertical centering with bold formatting, and saves the file as HeaderMerged.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells merge cells C# | center text merged cell Aspose | bold header Excel Aspose.Cells | C# Excel styling Aspose | .NET Excel header merge
// Common Searches: Aspose.Cells merge range and center text C# | how to create a centered header in Excel with Aspose.Cells | C# Aspose.Cells set horizontal vertical alignment | merge cells A2:D2 Aspose.Cells example
// Developer Intent: Generate a single header cell spanning A2:D2, assign text, and apply centered bold styling in a .NET workbook.
// Use Cases: Building a report template with a title row that spans multiple columns. | Designing an invoice sheet where the invoice title occupies A2:D2 and appears centered and bold. | Creating a dashboard worksheet with a merged header for section headings.
// AI Prompts: Provide C# code using Aspose.Cells to merge A2:D2, set "Header", and apply centered bold formatting. | Show how to style a merged header cell horizontally and vertically in Aspose.Cells for .NET. | Explain reusing a Style object for multiple merged header cells in an Aspose.Cells workbook.

using Aspose.Cells;

// Creates a new workbook, merges cells A2:D2 on the first worksheet, sets the value "Header", applies horizontal and vertical centering with bold formatting, and saves the file as HeaderMerged.xlsx using Aspose.Cells.
class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Merge cells A2:D2 (row index 1, column index 0, 1 row, 4 columns)
        cells.Merge(1, 0, 1, 4);

        // Set the header text in the merged cell
        cells[1, 0].Value = "Header";

        // Create a style that centers the text horizontally and vertically
        Style style = cells[1, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true; // optional: make the header bold
        cells[1, 0].SetStyle(style);

        // Save the workbook
        workbook.Save("HeaderMerged.xlsx");
    }
}
