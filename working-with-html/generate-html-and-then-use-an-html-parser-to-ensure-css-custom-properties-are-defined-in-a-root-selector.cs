// Title: C# – Save Aspose.Cells Workbook as HTML with :root CSS Custom Properties and Verify It
// Description: Creates a workbook, adds sample cells, configures HtmlSaveOptions to enable CSS custom properties, embeds a :root selector with a dummy variable, saves the workbook as a single HTML file, then reads the file and confirms that the :root selector and a CSS custom property (e.g., --demo‑color) exist.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | C# :root selector | CSS custom properties in Aspose.Cells | SaveAsSingleFile HTML | verify CSS variable presence | embed CSS variables in workbook HTML | Aspose.Cells HtmlSaveOptions example
// Common Searches: how to enable CSS variables when saving Aspose.Cells to HTML C# | C# code to add :root selector with custom properties in Aspose.Cells HTML output | verify :root selector exists in generated HTML Aspose.Cells | Aspose.Cells HtmlSaveOptions SaveAsSingleFile custom CSS example | parse saved HTML to check for CSS custom property
// Developer Intent: Generate a single‑file HTML export from a workbook that includes a :root CSS custom property and programmatically confirm its presence.
// Use Cases: Export a spreadsheet to a self‑contained HTML page with theme‑able CSS variables. | Inject a custom :root selector (e.g., --demo-color) into the HTML output for consistent styling across the page. | Automate validation of the exported HTML to ensure required CSS custom properties are correctly embedded.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as a single HTML file that contains a :root selector defining CSS custom properties. | Provide a method to load the generated HTML and assert that a :root selector with a specific CSS variable (e.g., --demo-color) is present. | Suggest robust error‑handling for missing HTML output or absent :root selector when validating Aspose.Cells HTML exports.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlCssCustomPropertiesDemo
{
    // Creates a workbook, adds sample cells, configures HtmlSaveOptions to enable CSS custom properties, embeds a :root selector with a dummy variable, saves the workbook as a single HTML file, then reads the file and confirms that the :root selector and a CSS custom property (e.g., --demo‑color) exist.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add some sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // -----------------------------------------------------------------
                // 2. Configure HtmlSaveOptions.
                //    - Enable CSS custom properties to allow reuse of resources.
                //    - Save as a single file so that CSS can be embedded.
                //    - Add a custom :root selector with a dummy CSS custom property.
                // -----------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true,
                    SaveAsSingleFile = true,
                    CssStyles = @"
                        :root {
                            --demo-color: #ff6600;
                        }
                        body {
                            color: var(--demo-color);
                        }"
                };

                // -----------------------------------------------------------------
                // 3. Save the workbook as HTML.
                // -----------------------------------------------------------------
                string htmlPath = "output.html";
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML at '{htmlPath}'.");

                // -----------------------------------------------------------------
                // 4. Load the generated HTML and verify that a :root selector
                //    containing a CSS custom property exists.
                // -----------------------------------------------------------------
                bool rootFound = false;

                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    // Simple verification: check for ":root" and a CSS custom property prefix "--"
                    if (htmlContent.Contains(":root") && htmlContent.Contains("--"))
                    {
                        rootFound = true;
                    }
                }
                else
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not found.");
                }

                Console.WriteLine(rootFound
                    ? "Verification succeeded: :root selector with CSS custom properties is present."
                    : "Verification failed: :root selector with CSS custom properties not found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
