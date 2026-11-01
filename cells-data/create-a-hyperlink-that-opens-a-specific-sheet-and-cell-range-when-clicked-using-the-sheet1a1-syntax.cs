// Title: How to add an internal '#Sheet1!A1' hyperlink in another worksheet cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that inserts a hyperlink in Sheet2!B2 pointing to Sheet1!A1 using the '#Sheet1!A1' address. | Show how to apply blue font color and single underline styling to the hyperlink cell with Aspose.Cells styling API. | Demonstrate saving the workbook as an .xlsx file after creating the internal sheet hyperlink.
// Common Searches: Aspose.Cells C# internal hyperlink to another sheet using #Sheet1!A1 syntax | Create a link from Sheet2 to Sheet1!A1 in Aspose.Cells .NET | C# Aspose.Cells format hyperlink with blue underline | Save workbook with cross‑sheet hyperlink using Aspose.Cells for .NET
// Tags: Aspose.Cells create cross‑sheet hyperlink C# | hyperlink styling blue underline Aspose.Cells | save workbook as xlsx Aspose.Cells .NET | worksheet hyperlink address #Sheet1!A1 | add internal sheet link Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, renames the first worksheet to Sheet1, writes a value to A1, adds a second worksheet named Sheet2, places the text "Go to Sheet1!A1" in B2, inserts an internal hyperlink '#Sheet1!A1' on that cell, styles it with blue underlined font, and saves the file as HyperlinkToSheet1.xlsx.
class HyperlinkDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the default first worksheet and rename it to "Sheet1"
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Put some sample data in Sheet1!A1
            sheet1.Cells["A1"].PutValue("This is Sheet1!A1");

            // Add a second worksheet where the hyperlink will be placed
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Set display text for the hyperlink in Sheet2!B2
            sheet2.Cells["B2"].PutValue("Go to Sheet1!A1");

            // Add a hyperlink that uses the '#Sheet1!A1' syntax
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlinkAddress
            sheet2.Hyperlinks.Add(1, 1, 0, 0, "#Sheet1!A1"); // B2 cell (row index 1, column index 1)

            // Apply typical hyperlink styling (blue and underlined)
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = Color.Blue;
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single;
            sheet2.Cells["B2"].SetStyle(hyperlinkStyle);

            // Save the workbook to a file
            workbook.Save("HyperlinkToSheet1.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
