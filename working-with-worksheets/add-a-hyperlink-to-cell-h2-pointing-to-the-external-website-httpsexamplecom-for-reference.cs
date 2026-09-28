// Title: How to add an external hyperlink to cell H2 in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, puts the text "Example" in cell H2, and attaches a hyperlink to https://example.com using Aspose.Cells. | Show the exact Aspose.Cells Hyperlinks.Add call needed to link cell H2 (row 1, column 7) to an external website in a .NET application. | Provide a complete C# snippet that saves an Excel file with a clickable link in H2 pointing to https://example.com.
// Common Searches: Aspose.Cells C# add hyperlink to a specific cell | How to create an external URL link in Excel using Aspose.Cells .NET | C# Hyperlinks.Add method example for linking a cell to a website
// Tags: Aspose.Cells Hyperlinks.Add external URL | C# add hyperlink to Excel cell | hyperlink Excel cell using Aspose.Cells | Excel file with clickable link Aspose.Cells | link external website in workbook Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes "Example" into cell H2, adds a hyperlink that points to https://example.com on that cell using the Hyperlinks.Add method, and saves the file as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set display text in cell H2
        sheet.Cells["H2"].PutValue("Example");

        // Add a hyperlink to cell H2 (row index 1, column index 7) pointing to the external website
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink
        sheet.Hyperlinks.Add(1, 7, 1, 1, "https://example.com");

        // Save the workbook
        workbook.Save("HyperlinkExample.xlsx");
    }
}
