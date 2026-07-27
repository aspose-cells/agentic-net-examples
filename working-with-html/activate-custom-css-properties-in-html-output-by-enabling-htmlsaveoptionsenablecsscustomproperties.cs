// Title: How to Enable CSS Custom Properties (Variables) in HTML Export with Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, adding data, configuring HtmlSaveOptions.EnableCssCustomProperties, and saving the file as HTML that contains CSS variables for modern web styling.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | EnableCssCustomProperties | CSS variables | custom properties | HTML export | Excel to HTML | web styling
// Common Searches: Aspose.Cells enable CSS variables in HTML | HtmlSaveOptions EnableCssCustomProperties example C# | Export Excel to HTML with CSS custom properties | How to use EnableCssCustomProperties in Aspose.Cells | Save workbook as HTML with CSS variables .NET
// Developer Intent: Activate CSS custom properties when converting an Excel workbook to HTML using Aspose.Cells.
// Use Cases: Generate HTML reports that use CSS variables for easy theming and brand updates. | Create web pages from Excel data where colors, fonts, or spacing can be adjusted via CSS custom properties. | Produce HTML output compatible with browsers that support CSS variables for responsive and dynamic designs.
// AI Prompts: Show me a C# example that saves an Excel workbook to HTML with EnableCssCustomProperties set to true. | Explain how to modify the generated CSS variable names after exporting a workbook with Aspose.Cells. | Provide a step‑by‑step guide to enable CSS custom properties in Aspose.Cells HTML export and apply a dark‑mode theme using those variables.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding data, configuring HtmlSaveOptions.EnableCssCustomProperties, and saving the file as HTML that contains CSS variables for modern web styling.
    public class EnableCssCustomPropertiesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to the worksheet
                worksheet.Cells["A1"].PutValue("Hello");
                worksheet.Cells["A2"].PutValue("World");

                // Initialize HTML save options and enable CSS custom properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true // Activate custom CSS properties
                };

                // Save the workbook as an HTML file using the configured options
                string outputPath = "HtmlWithCustomProperties.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableCssCustomPropertiesDemo.Run();
        }
    }
}
