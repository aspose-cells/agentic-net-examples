// Title: Update a ListObject's TableStyleName after exporting to HTML with Aspose.Cells for .NET and confirm the new CSS class
// AI Prompts: Modify the TableStyleName of the first ListObject in a workbook, re‑save the workbook as HTML, and check the generated file for the new class attribute using C# and Aspose.Cells. | Write C# code that loads an Excel file, exports it to HTML, changes the table style name, exports again, and validates that the HTML contains the updated CSS class.
// Common Searches: aspocells change ListObject TableStyleName after HTML export c# | how to verify updated table CSS class in Aspose.Cells generated HTML | re‑export workbook to HTML after modifying table style with Aspose.Cells | c# Aspose.Cells update table style name and check HTML output | detect new class attribute in HTML produced by Aspose.Cells after style change
// Tags: Aspose.Cells ListObject TableStyleName update | HTML export after table style change | verify CSS class in Aspose.Cells HTML output | C# re‑export workbook to HTML | programmatic table style modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The example loads an Excel workbook, saves it as HTML, changes the TableStyleName of the first ListObject, re‑saves to a new HTML file, and reads the output to confirm that the new CSS class appears in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string initialHtmlPath = "output_initial.html";
            const string updatedHtmlPath = "output_updated.html";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Initial export to HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            workbook.Save(initialHtmlPath, htmlOptions);

            // Change the style name of the first table (ListObject) in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.ListObjects.Count > 0)
            {
                // Modify the first table
                ListObject table = sheet.ListObjects[0];
                // Use TableStyleName to affect the generated HTML class attribute
                table.TableStyleName = "newTableClass";
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Re‑export the workbook to HTML after the change
            workbook.Save(updatedHtmlPath, htmlOptions);

            // Verify that the updated HTML contains the new class name
            if (!File.Exists(updatedHtmlPath))
            {
                Console.WriteLine($"Updated HTML file \"{updatedHtmlPath}\" was not created.");
                return;
            }

            string updatedHtml = File.ReadAllText(updatedHtmlPath);
            if (updatedHtml.Contains("class=\"newTableClass\""))
            {
                Console.WriteLine("Verification succeeded: Table style name was reflected in the HTML.");
            }
            else
            {
                Console.WriteLine("Verification failed: Updated class name not found in the HTML.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
