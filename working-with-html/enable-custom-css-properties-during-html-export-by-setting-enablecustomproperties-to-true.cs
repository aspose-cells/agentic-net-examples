// Title: C# – Export Excel to HTML with CSS Custom Properties (EnableCssCustomProperties) using Aspose.Cells
// Description: Shows how to create a workbook, enable HtmlSaveOptions.EnableCssCustomProperties, and save the file as HTML. The option consolidates repeated styles into CSS variables, shrinking the markup and allowing centralized styling in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | EnableCssCustomProperties | CSS custom properties | HTML export | Excel to HTML | CSS variables | optimize HTML size
// Common Searches: Aspose.Cells EnableCssCustomProperties C# example | How to export Excel as HTML with CSS variables using Aspose | HtmlSaveOptions EnableCssCustomProperties sample code | Reduce HTML size Aspose.Cells export | C# export workbook to HTML with CSS custom properties
// Developer Intent: Activate CSS custom properties when saving a workbook as HTML with Aspose.Cells.
// Use Cases: Minimize generated HTML by converting repeated style values into CSS variables. | Produce web‑ready reports where theme colors and fonts are controlled centrally via CSS custom properties. | Maintain consistent styling across multiple HTML pages exported from different workbooks.
// AI Prompts: Provide a C# snippet that exports an Aspose.Cells workbook to HTML with EnableCssCustomProperties set to true and links an external stylesheet. | Show code that toggles HtmlSaveOptions.EnableCssCustomProperties based on a configuration setting in a .NET application. | Explain how EnableCssCustomProperties changes the HTML output and how to reference the generated CSS variables in external CSS.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, enable HtmlSaveOptions.EnableCssCustomProperties, and save the file as HTML. The option consolidates repeated styles into CSS variables, shrinking the markup and allowing centralized styling in Aspose.Cells for .NET.
    public class EnableCssCustomPropertiesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello Aspose.Cells");

                // Create HTML save options and enable CSS custom properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
                htmlOptions.EnableCssCustomProperties = true; // Optimize HTML using CSS custom properties

                // Save the workbook as HTML with the custom property optimization enabled
                string outputPath = "HtmlWithCustomProperties.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableCssCustomPropertiesDemo.Run();
        }
    }
}
