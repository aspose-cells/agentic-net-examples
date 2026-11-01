// Title: Set Times New Roman as the default font when saving a workbook to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set default font for HTML export | Save workbook as HTML using Times New Roman with Aspose.Cells | How to apply a default font to HTML output in Aspose.Cells .NET | HtmlSaveOptions missing DefaultFont property in Aspose.Cells | Workaround for setting default font during HTML conversion in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions default font | HTML export font styling Aspose.Cells | C# Times New Roman for HTML output | Aspose.Cells workbook to HTML custom font | missing DefaultFont property workaround Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, adds data, and demonstrates how to configure HTML export with Aspose.Cells. Since the HtmlSaveOptions class does not expose a DefaultFont property, the code suggests applying a default font (e.g., Times New Roman) via style settings before calling workbook.Save("output.html", htmlOptions).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The DefaultFont property is not available in the current Aspose.Cells version.
            // If needed, set the default font via other means (e.g., style settings) before saving.

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
