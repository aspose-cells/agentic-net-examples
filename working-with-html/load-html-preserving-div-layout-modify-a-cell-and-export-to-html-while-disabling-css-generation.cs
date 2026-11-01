// Title: Load an HTML file with DIV layout into Aspose.Cells, modify cell B2, and save as HTML without generating CSS (C#)
// AI Prompts: Read an HTML document into a Workbook, change the value of cell B2 on the first worksheet, and export the active sheet to HTML while disabling CSS stylesheet creation using Aspose.Cells for .NET. | Using Aspose.Cells, load an HTML file preserving its DIV structure, update a specific cell, and save the workbook as HTML with CssStyleSheetType set to None to suppress CSS output.
// Common Searches: asp.net load html into Aspose.Cells workbook keeping div tags | change cell content after loading html with Aspose.Cells C# example | save only the current worksheet as html without css using Aspose.Cells | how to turn off css stylesheet generation in Aspose.Cells html export | set CssStyleSheetType to None in HtmlSaveOptions Aspose.Cells
// Tags: htmlloadoptions preserve div layout | update cell value Aspose.Cells C# | htmlsaveoptions no css stylesheet | export active worksheet only html | disable css generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook while retaining the original DIV layout, updates cell B2 in the first worksheet, and saves the workbook as HTML. It exports only the active worksheet and suppresses CSS generation by setting the CssStyleSheetType to None when the property is available.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.html";

            // Ensure the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file with default options
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Modify cell B2 in the first worksheet
            var cell = workbook.Worksheets[0].Cells["B2"];
            cell.PutValue("New Value");

            // Prepare save options
            var saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true // export only the active sheet
            };

            // If the CssStyleSheetType property exists, set it to None to suppress CSS generation
            var cssProp = typeof(HtmlSaveOptions).GetProperty("CssStyleSheetType");
            if (cssProp != null && cssProp.CanWrite)
            {
                var enumType = cssProp.PropertyType;
                var noneValue = Enum.Parse(enumType, "None");
                cssProp.SetValue(saveOptions, noneValue);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
