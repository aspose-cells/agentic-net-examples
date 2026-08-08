// Title: Export a Workbook with Merged Cells to One HTML File using Aspose.Cells for .NET
// Description: C# example that creates a workbook with two worksheets, merges a range on the second sheet, and saves the entire workbook as a single HTML page. The HtmlSaveOptions are set to SaveAsSingleFile and ShowAllSheets to preserve merged cell formatting and display all sheets in one file.
// Keywords: Aspose.Cells HTML export | C# save workbook as single HTML | merged cells HTML Aspose | ShowAllSheets property | SaveAsSingleFile example | convert Excel to HTML .NET | Aspose.Cells workbook merging
// Common Searches: Aspose.Cells export merged cells to HTML | C# save multiple worksheets as one HTML page | HtmlSaveOptions ShowAllSheets usage | How to create single HTML file from Excel with Aspose | Preserve merged regions when converting Excel to HTML
// Developer Intent: Generate a single HTML document that renders all worksheets and merged cells from an Aspose.Cells workbook.
// Use Cases: Web‑based preview of Excel reports without requiring Office installed | Single‑page HTML report for email or intranet distribution | Embedding Excel data with merged formatting into a web application
// AI Prompts: Write C# code with Aspose.Cells to export a multi‑sheet workbook containing merged cells to one HTML file. | Explain the impact of HtmlSaveOptions.ShowAllSheets and SaveAsSingleFile on the HTML output for merged regions. | Provide troubleshooting steps when merged cells lose their formatting after HTML conversion.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a workbook with two worksheets, merges a range on the second sheet, and saves the entire workbook as a single HTML page. The HtmlSaveOptions are set to SaveAsSingleFile and ShowAllSheets to preserve merged cell formatting and display all sheets in one file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Content of the first sheet");

        // Add a second worksheet and fill it with data
        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["B2"].PutValue("Content of the second sheet");

        // Demonstrate a merged region in the second sheet
        // Merge cells B2:D3 and set a value for the merged area
        sheet2.Cells.Merge(1, 1, 2, 3); // Row 1, Column 1 (B2) spanning 2 rows and 3 columns
        sheet2.Cells["B2"].PutValue("Merged Area");

        // Configure HTML save options:
        // - Save as a single HTML file
        // - Show all worksheets within that single file
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            SaveAsSingleFile = true,
            ShowAllSheets = true
        };

        // Define the output HTML file path
        string outputPath = "MergedWorkbook.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        // Inform the user where the file was saved
        Console.WriteLine($"HTML representation saved to: {Path.GetFullPath(outputPath)}");
    }
}
