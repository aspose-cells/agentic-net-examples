// Title: Create a mailto: email hyperlink in an Excel cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add mailto hyperlink to a specific cell | How to generate a clickable email link in Excel using Aspose.Cells .NET | Set cell style and email hyperlink programmatically with Aspose.Cells | Create Excel workbook with email hyperlink and custom subject using C# | Add mailto link to A1 cell using Aspose.Cells Hyperlinks.Add
// Tags: Aspose.Cells Hyperlinks.Add mailto | C# create email hyperlink Excel | Aspose.Cells style blue underlined hyperlink | Save workbook as Xlsx with email link | Excel cell mailto link Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    // The example creates a new workbook, writes "Email Me" in cell A1, adds a mailto: hyperlink with predefined subject and body, applies blue underlined hyperlink styling, and saves the file as MailtoHyperlink.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the display text for the cell
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Email Me");

                // Define the mailto hyperlink (customize as needed)
                string mailtoLink = "mailto:someone@example.com?subject=Hello%20World&body=Please%20enter%20your%20message%20here";

                // Add the hyperlink to cell A1 (row 0, column 0)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, url
                sheet.Hyperlinks.Add(0, 0, 1, 1, mailtoLink);

                // Format the cell to look like a hyperlink (blue and underlined)
                Style style = workbook.CreateStyle();
                style.Font.Color = Color.Blue;
                style.Font.Underline = FontUnderlineType.Single;

                // Apply the style to the cell
                cell.SetStyle(style);

                // Save the workbook to a file
                string outputPath = "MailtoHyperlink.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
