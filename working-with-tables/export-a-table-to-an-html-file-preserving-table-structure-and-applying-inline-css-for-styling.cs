// Title: Export an Excel ListObject table to HTML with inline CSS and Base64‑encoded images using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a worksheet, adds a ListObject table, applies a built‑in table style, and saves the workbook as an HTML file with inline CSS and images embedded as Base64 using Aspose.Cells. | Show how to configure HtmlSaveOptions in Aspose.Cells to preserve the table structure and embed CSS directly in the generated HTML output. | Demonstrate customizing the inline CSS for the exported HTML table while keeping the ListObject formatting intact.
// Common Searches: how to save an Excel ListObject as HTML with inline styles using Aspose.Cells C# | Aspose.Cells export worksheet to HTML preserving table formatting and embedding images as Base64 | C# example for converting a range to an HTML table with embedded CSS via Aspose.Cells | customize HTML output CSS for Excel tables in Aspose.Cells .NET | export Excel table to HTML file with built‑in table style Aspose.Cells
// Tags: export ListObject to HTML Aspose.Cells | inline CSS for HTML export Aspose.Cells | HtmlSaveOptions Base64 image embedding | apply built‑in table style during HTML conversion | convert Excel range to HTML table C#

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program creates a workbook, populates sample data, converts a range into a ListObject table, applies a built‑in style, and saves the worksheet as an HTML file with inline CSS and Base64‑encoded images using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Convert the range into an Excel table (ListObject)
            // Add returns the index of the created table; retrieve the ListObject via the collection
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Optional: configure table appearance
            table.ShowHeaderRow = true;               // show header row
            table.TableStyleType = TableStyleType.TableStyleMedium2; // apply built‑in style

            // Configure HTML save options to embed images as Base64
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save("ExportedTable.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
