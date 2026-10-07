// Title: How to change TableCssId for an HTML export with Aspose.Cells in C# and verify the new CSS class
// AI Prompts: Generate C# code that loads an Excel workbook, saves it to HTML, sets HtmlSaveOptions.TableCssId to a custom value, re‑saves the file, and checks the output HTML for the new class attribute. | Show how to programmatically update the table CSS identifier after an initial HTML export using Aspose.Cells and confirm the change by reading the saved HTML file.
// Common Searches: Aspose.Cells C# change TableCssId after first HTML export | verify custom table CSS class in HTML saved by Aspose.Cells | re‑export workbook with new TableCssId using HtmlSaveOptions | how to set custom table class name in Aspose.Cells HTML output C#
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | custom table CSS class Aspose.Cells | re‑export workbook with modified HTML options | validate HTML class attribute C# | programmatic HTML export class name update

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, exports it to HTML with default settings, then changes HtmlSaveOptions.TableCssId to "myTableClass", re‑exports the workbook, reads the second HTML file, and confirms that the generated HTML contains the new class attribute.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // First export with default TableCssId
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        workbook.Save("output1.html", htmlOptions);

        // Change the TableCssId to a new class name
        htmlOptions.TableCssId = "myTableClass";

        // Re‑export the workbook with the updated TableCssId
        workbook.Save("output2.html", htmlOptions);

        // Verify that the new class name appears in the generated HTML
        string htmlContent = File.ReadAllText("output2.html");
        bool classFound = htmlContent.Contains("class=\"myTableClass\"");
        Console.WriteLine($"Updated class present: {classFound}");
    }
}
