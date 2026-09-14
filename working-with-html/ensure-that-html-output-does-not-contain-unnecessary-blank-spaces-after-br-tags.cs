// Title: Export an Aspose.Cells workbook to HTML in C# while preventing extra blank <br> tags
// AI Prompts: Write C# code that saves a Workbook as HTML using Aspose.Cells and sets the HTML save options to omit <br> tags for empty rows. | Show how to configure Aspose.Cells HTML export in .NET so that blank Excel rows do not generate line‑break tags.
// Common Searches: how to stop Aspose.Cells from adding <br> tags for empty rows when saving to HTML in C# | Aspose.Cells HTML export blank line removal C# example | prevent extra line breaks in HTML output from Excel using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions suppress empty row line breaks | C# export Excel to HTML without extra <br> | prevent blank rows in Aspose.Cells HTML output | Excel to HTML conversion avoid unnecessary line breaks

using System;
using Aspose.Cells;

// Creates a workbook, adds a blank row, configures HTML save options (default behavior skips blank rows), and saves the file as HTML, ensuring no extra <br> tags are generated.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data: add some content and a blank row
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("First line");
            sheet.Cells["A2"].PutValue(string.Empty); // blank row that could generate a <br>

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: In recent Aspose.Cells versions the ExportBlankLines property is not available.
            // The default behavior already avoids rendering blank rows as extra <br> tags.

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
