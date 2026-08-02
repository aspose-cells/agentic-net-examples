// Title: C# – Merge Cells A1:D1 into a Bold, Centered Header Row with Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, merge the range A1‑D1 on the first worksheet, insert custom header text, apply bold font with center alignment, and save the file as HeaderMerged.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | merge cells | Excel header row | bold font | center alignment | style formatting | workbook automation | worksheet styling
// Common Searches: Aspose.Cells merge cells A1 D1 C# | Create a merged header row in Excel with Aspose.Cells | Apply bold and centered style to merged cells using .NET | C# code to add a title across columns A‑D in Aspose.Cells | How to format a merged header in Aspose.Cells for .NET
// Developer Intent: Generate a single‑row title that spans columns A‑D, displays custom text, and uses bold, centered formatting.
// Use Cases: Building a financial report where the report title stretches across the first four columns. | Designing an invoice template with a prominent, centered heading that covers A‑D. | Creating a dashboard sheet that groups sections with bold, merged headers for visual clarity.
// AI Prompts: Provide C# code that merges A1:D1, writes a user‑defined header, makes the text bold and centered, then saves the workbook with Aspose.Cells. | Show how to define a reusable style for merged title rows, including font size, color, and background, and apply it to multiple worksheets. | Explain step‑by‑step how to programmatically create a merged header row and customize its appearance in Aspose.Cells for .NET.

using Aspose.Cells;

// Demonstrates how to create a new workbook, merge the range A1‑D1 on the first worksheet, insert custom header text, apply bold font with center alignment, and save the file as HeaderMerged.xlsx using Aspose.Cells for .NET.
class HeaderMergeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells A1:D1 (row 0, columns 0 to 3)
        cells.Merge(0, 0, 1, 4);

        // Set the header text in the merged cell
        cells[0, 0].PutValue("Header Row");

        // Apply bold formatting (and center alignment) to the merged cell
        Style style = cells[0, 0].GetStyle();
        style.Font.IsBold = true;
        style.HorizontalAlignment = TextAlignmentType.Center;
        cells[0, 0].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("HeaderMerged.xlsx");
    }
}
