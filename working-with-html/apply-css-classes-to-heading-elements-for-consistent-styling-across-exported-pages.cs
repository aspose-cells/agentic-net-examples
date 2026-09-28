// Title: Apply a CSS class to heading cells when exporting an Excel worksheet to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, assigns a CSS class named "heading" to each cell in the first row, and saves the workbook as HTML using Aspose.Cells. | Show how to use the Style.Custom property to embed a CSS class name for heading cells during HTML export with Aspose.Cells. | Demonstrate setting HtmlSaveOptions so that only the current worksheet is saved as HTML and the heading cells retain their assigned CSS class.
// Common Searches: Aspose.Cells C# export first row as HTML with custom CSS class | How to add a CSS class to Excel cells for HTML output using Aspose.Cells .NET | Set custom style property for heading cells when saving workbook to HTML in C# | Export only active worksheet to HTML with Aspose.Cells while preserving CSS classes
// Tags: apply css class to heading cells Aspose.Cells | style.custom property html export Aspose.Cells | htmlsaveoptions exportactiveworksheetonly C# | excel to html custom css class Aspose.Cells | first row styling during workbook to html conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, sets the Style.Custom property of each cell in the first row to "class:heading", configures HtmlSaveOptions to export only the active worksheet, and saves the workbook as an HTML file where the heading cells carry the specified CSS class.
class ApplyCssToHeadings
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // CSS class name to be applied to heading cells
            const string headingCssClass = "heading";

            // Apply the CSS class to all used cells in the first row
            int lastColumn = sheet.Cells.MaxColumn;
            for (int col = 0; col <= lastColumn; col++)
            {
                Cell cell = sheet.Cells[0, col];
                Style style = cell.GetStyle();

                // Set custom CSS class for HTML export (format: "class:YourClassName")
                style.Custom = $"class:{headingCssClass}";
                cell.SetStyle(style);
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export only the active worksheet (optional)
                ExportActiveWorksheetOnly = true
                // Note: ExportCssClassNames property is not required for this version
            };

            // Save the workbook as an HTML file; heading cells will carry the specified CSS class
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
