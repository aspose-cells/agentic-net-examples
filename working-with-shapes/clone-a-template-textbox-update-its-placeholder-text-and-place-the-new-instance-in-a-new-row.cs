// Title: How to clone a TextBox shape, update its placeholder text, and place it in a new worksheet row with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells for C# to duplicate a TextBox named "TemplateBox", set its Text property to a custom placeholder, and add the cloned shape to the row after the last data row. | Write C# code that loads an Excel workbook, finds a TextBox shape, clones it with AddCopy, changes the cloned TextBox's text, and saves the workbook.
// Common Searches: Aspose.Cells C# clone TextBox shape and change its text | AddCopy shape to specific row in Excel using Aspose.Cells | Copy a TextBox from a template worksheet and insert into new row with Aspose.Cells | How to update placeholder text of a cloned TextBox in Aspose.Cells | C# Aspose.Cells find shape by name and duplicate it
// Tags: clone textbox shape Aspose.Cells C# | addcopy shape to worksheet row | update textbox placeholder Aspose.Cells | find shape by name Aspose.Cells | duplicate template shape Excel C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, locates a TextBox named "TemplateBox", clones it to a row after the existing data using AddCopy, updates the cloned TextBox's placeholder text, and saves the result, with comprehensive error handling for missing files and shape operations.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(templatePath);
            var sheet = workbook.Worksheets[0];

            // Locate the template TextBox (assumed name "TemplateBox")
            Shape templateShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape is TextBox tb && tb.Name == "TemplateBox")
                {
                    templateShape = tb; // store as Shape to access position properties
                    break;
                }
            }

            if (templateShape == null)
            {
                Console.WriteLine("Template TextBox \"TemplateBox\" not found.");
                return;
            }

            // Determine the row where the new TextBox will be placed
            int targetRow = sheet.Cells.MaxDataRow + 2; // one row after existing data

            // Clone the template TextBox onto the worksheet using AddCopy with offsets
            Shape clonedShape;
            try
            {
                clonedShape = sheet.Shapes.AddCopy(
                    templateShape,
                    targetRow,
                    templateShape.UpperLeftColumn,
                    0, // UpperLeftRowOffset (not required for this example)
                    0  // UpperLeftColumnOffset (not required for this example)
                );
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Failed to clone shape: {copyEx.Message}");
                return;
            }

            // Cast the cloned shape back to TextBox and update its text
            if (clonedShape is TextBox newBox)
            {
                newBox.Text = "Updated placeholder text";
            }
            else
            {
                Console.WriteLine("Cloned shape is not a TextBox.");
                return;
            }

            // Save the modified workbook
            try
            {
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to \"{resultPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
