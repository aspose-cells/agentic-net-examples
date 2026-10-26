// Title: Add an external website hyperlink to cell H5 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells Hyperlinks.Add example for inserting a website link into cell H5 | How to add an external URL hyperlink to a specific Excel cell with Aspose.Cells .NET | Insert clickable link to https://www.example.com in Excel using Aspose.Cells C# API | Aspose.Cells tutorial for adding a hyperlink to a particular cell in a workbook | Create an Excel hyperlink in cell H5 programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells Hyperlinks.Add external link | C# insert Excel hyperlink with Aspose.Cells | Add website URL to specific cell using .NET Excel API | Programmatic Excel hyperlink creation Aspose.Cells | Cell H5 hyperlink Aspose.Cells example

using Aspose.Cells;

// The sample creates a new workbook, accesses the first worksheet, and uses the Hyperlinks.Add method to place an external URL (https://www.example.com) into cell H5 (row 5, column H). The workbook is then saved as HyperlinkDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H5 (row 5, column H)
        // Row and column indices are zero‑based, so H5 is (4, 7)
        // The hyperlink points to an external website
        sheet.Hyperlinks.Add(4, 7, 1, 1, "https://www.example.com");

        // Save the workbook
        workbook.Save("HyperlinkDemo.xlsx");
    }
}
