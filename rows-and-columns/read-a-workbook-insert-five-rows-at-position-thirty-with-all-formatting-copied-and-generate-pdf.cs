// Title: Insert five rows at row 30 with formatting preserved and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert multiple rows at a specific index while keeping cell styles | convert modified Excel worksheet to PDF after inserting rows with Aspose.Cells | C# code to add rows at row 30 and export workbook to PDF using Aspose.Cells | preserve formatting when inserting rows in an Aspose.Cells workbook | Aspose.Cells insert rows and generate PDF in a single operation
// Tags: Aspose.Cells InsertRows preserving formatting | C# insert rows at specific worksheet index | Aspose.Cells export worksheet to PDF | SaveFormat.Pdf parameter in Aspose.Cells Save method | copy cell styles during row insertion Aspose.Cells

using Aspose.Cells;

// The example loads an existing Excel file (input.xlsx), inserts five rows at zero‑based index 29 (row 30) on the first worksheet while copying the original formatting, and then saves the updated workbook directly as a PDF file (output.pdf) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        var workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        var worksheet = workbook.Worksheets[0];

        // Insert five rows at position 30 (row index 29, zero‑based) and copy formatting
        worksheet.Cells.InsertRows(29, 5, true);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
