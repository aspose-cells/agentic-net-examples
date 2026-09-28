// Title: How to preserve worksheet name capitalization in HTML heading tags when exporting with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML and ensures the <h1> heading displays the worksheet name with its original capitalization. | Show how to configure HtmlSaveOptions in Aspose.Cells so that the generated HTML keeps the exact case of each worksheet title.
// Common Searches: Aspose.Cells export to HTML keep original worksheet name case | C# HtmlSaveOptions preserve sheet name capitalization in HTML output | How to retain worksheet title casing when saving workbook as HTML with Aspose.Cells | HTML heading tags reflect worksheet name case Aspose.Cells .NET | Save workbook to HTML without changing sheet name case using Aspose.Cells
// Tags: Aspose.Cells HTML export keep worksheet name case | HtmlSaveOptions heading tag case handling | C# save workbook as HTML original sheet title | maintain worksheet name case in generated HTML | Aspose.Cells heading tag case sensitivity

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, sets a worksheet name with specific capitalization, and saves the workbook as HTML using HtmlSaveOptions. The generated HTML heading tags (<h1>) retain the original case of the worksheet name.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet and set its name with specific capitalization
                var worksheet = workbook.Worksheets[0];
                worksheet.Name = "MyCustomSheetName"; // original capitalization preserved

                // Add some sample data (optional)
                worksheet.Cells["A1"].PutValue("Sample Data");

                // Configure HTML save options
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export all worksheets (set to true if you want the active sheet only)
                    ExportActiveWorksheetOnly = false
                };

                // Save the workbook as HTML; the heading tags will use the original worksheet name capitalization
                workbook.Save("Output.html", htmlOptions);
                Console.WriteLine("Workbook saved successfully as Output.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
