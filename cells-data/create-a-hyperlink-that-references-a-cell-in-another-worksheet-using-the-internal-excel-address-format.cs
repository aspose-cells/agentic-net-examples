// Title: Create an internal worksheet hyperlink to a specific cell using Aspose.Cells for .NET (C#)
// Common Searches: how to add a hyperlink from Sheet1!A1 to Sheet2!B2 with Aspose.Cells C# | Aspose.Cells C# example for internal Excel address format hyperlink | using Aspose.Cells Hyperlinks.Add to navigate between worksheets | C# code to link a cell to another worksheet cell in an Excel file with Aspose.Cells | Aspose.Cells create worksheet‑to‑worksheet hyperlink programmatically
// Tags: Aspose.Cells Hyperlinks.Add internal address | C# worksheet-to-worksheet hyperlink | Excel internal address format Aspose.Cells | Aspose.Cells hyperlink to specific cell | save workbook with hyperlink Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, renames the first sheet to "Sheet1", adds a second sheet named "Sheet2", writes a value to Sheet2!B2, and then inserts a hyperlink in Sheet1!A1 that points to Sheet2!B2 using the Hyperlinks.Add method with the internal Excel address format. The workbook is saved as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default) and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Put some data in Sheet2!B2 for demonstration
            sheet2.Cells["B2"].PutValue("Target Cell");

            // Add a hyperlink that points to Sheet2!B2
            // The Add method accepts the address as a string
            sheet1.Hyperlinks.Add(0, 0, 1, 1, "Sheet2!B2");

            // Set display text for the hyperlink (cell value)
            sheet1.Cells["A1"].PutValue("Go to Sheet2 B2");

            // Save the workbook
            workbook.Save("HyperlinkExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
