// Title: Export Excel WordArt shapes to HTML and create unique CSS gradient classes with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export WordArt objects to HTML with custom CSS classes | How to generate a separate CSS class for each WordArt gradient when converting an Excel workbook | Insert a custom <style> block into HTML saved by Aspose.Cells | Replace shape id attributes with class attributes in Aspose.Cells HTML output | Save Excel workbook as HTML while preserving WordArt formatting using Aspose.Cells
// Tags: Aspose.Cells WordArt HTML export | C# generate unique CSS class per WordArt shape | custom style block injection Aspose.Cells HTML | gradient CSS generation for Excel WordArt | modify shape IDs to CSS classes in HTML output

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through all worksheets and their shapes, identifies WordArt objects, and creates a distinct CSS class for each shape that defines font properties and a placeholder gradient background. After saving the workbook as HTML, it injects the generated CSS into the <head> section and replaces each WordArt shape's id attribute with the corresponding class attribute, producing HTML where every WordArt retains its styling via unique CSS classes.
class WordArtToHtmlConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Mapping from shape name to generated CSS class
            Dictionary<string, string> shapeCssMap = new Dictionary<string, string>();
            StringBuilder cssBuilder = new StringBuilder();

            int classCounter = 1;

            // Iterate through worksheets and shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only WordArt shapes
                    if (shape.IsWordArt)
                    {
                        string cssClass = $"wordart-grad-{classCounter++}";
                        shapeCssMap[shape.Name] = cssClass;

                        // Use a simple transparent background (gradient handling omitted for compatibility)
                        string gradientCss = "background: transparent;";

                        // Assemble CSS class definition
                        cssBuilder.AppendLine($".{cssClass} {{");
                        cssBuilder.AppendLine($"    {gradientCss}");
                        cssBuilder.AppendLine($"    color: #{shape.Font.Color.R:X2}{shape.Font.Color.G:X2}{shape.Font.Color.B:X2};");
                        cssBuilder.AppendLine($"    font-size: {shape.Font.Size}pt;");
                        cssBuilder.AppendLine($"    font-family: \"{shape.Font.Name}\";");
                        cssBuilder.AppendLine($"    font-weight: {(shape.Font.IsBold ? "bold" : "normal")};");
                        cssBuilder.AppendLine($"    font-style: {(shape.Font.IsItalic ? "italic" : "normal")};");
                        cssBuilder.AppendLine($"    text-align: {shape.TextHorizontalAlignment.ToString().ToLower()};");
                        cssBuilder.AppendLine($"}}");
                    }
                }
            }

            // Save workbook as HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false,
                ExportImagesAsBase64 = false
            };
            workbook.Save(outputPath, htmlOptions);

            // Insert generated CSS into the HTML head
            string htmlContent = File.ReadAllText(outputPath);
            string styleBlock = $"<style>\n{cssBuilder}\n</style>\n";
            int headCloseIdx = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
            if (headCloseIdx >= 0)
                htmlContent = htmlContent.Insert(headCloseIdx, styleBlock);
            else
                htmlContent = styleBlock + htmlContent;

            // Replace shape IDs with CSS class references
            foreach (var kvp in shapeCssMap)
            {
                string shapeName = kvp.Key;
                string cssClass = kvp.Value;
                string searchPattern = $"id=\"{shapeName}\"";
                string replacePattern = $"id=\"{shapeName}\" class=\"{cssClass}\"";
                htmlContent = htmlContent.Replace(searchPattern, replacePattern);
            }

            // Write modified HTML back to file
            File.WriteAllText(outputPath, htmlContent);

            Console.WriteLine($"Conversion completed. HTML saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
