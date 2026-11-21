// Title: Add an internal hyperlink to navigate from one worksheet cell to another using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# add hyperlink to another worksheet cell | how to create internal sheet navigation link with Aspose.Cells | C# code sample for worksheet hyperlink pointing to Sheet2 B2 | Aspose.Cells hyperlink collection internal reference example | navigate between worksheets using hyperlink in Aspose.Cells C#
// Tags: Aspose.Cells add internal worksheet hyperlink | C# hyperlink collection navigation between sheets | Aspose.Cells workbook internal link to cell | hyperlink to specific cell Aspose.Cells C# | Excel internal navigation hyperlink Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook with two worksheets, inserts an internal hyperlink in Sheet1!A1 that points to Sheet2!B2 using the Hyperlinks.Add method, sets display text, and saves the file as HyperlinkNavigation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (default created)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet to navigate to
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Example: place some text in the target cell on Sheet2
        sheet2.Cells["B2"].PutValue("Target Cell");

        // Add a hyperlink in Sheet1!A1 that points to Sheet2!B2
        // Hyperlink format for internal navigation: "Sheet2!B2"
        // The HyperlinkCollection.Add method parameters:
        // firstRow, firstColumn, totalRows, totalColumns, hyperlink
        sheet1.Hyperlinks.Add(0, 0, 1, 1, "Sheet2!B2");

        // Optional: set display text for the hyperlink
        sheet1.Cells["A1"].PutValue("Go to Sheet2!B2");

        // Save the workbook to a file
        workbook.Save("HyperlinkNavigation.xlsx");
    }
}
