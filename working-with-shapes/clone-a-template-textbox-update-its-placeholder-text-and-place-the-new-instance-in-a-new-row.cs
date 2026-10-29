// Title: Clone a TextBox shape, change its placeholder text, and insert the copy into a new row using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells clone a TextBox shape and set new placeholder text | Add a copied TextBox to the next empty row in an Excel worksheet with Aspose.Cells | How to use Shape.AddCopy to duplicate a TextBox and reposition it in a sheet using .NET | Aspose.Cells example for duplicating a TextBox and moving it to a new row | Update placeholder text of a cloned TextBox shape in C#
// Tags: duplicate TextBox shape using Aspose.Cells | set TextBox placeholder text C# | insert copied shape into next data row | Shape.AddCopy method example | reposition TextBox in Excel worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, finds a TextBox named "TemplateTextBox", clones it with Shape.AddCopy, updates the cloned TextBox's placeholder text, moves the clone to the row after the last used row while preserving the column, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the template TextBox shape by its name.
            TextBox? templateTextBox = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape is TextBox tb && tb.Name == "TemplateTextBox")
                {
                    templateTextBox = tb;
                    break;
                }
            }

            if (templateTextBox == null)
            {
                Console.WriteLine("Template TextBox not found.");
                return;
            }

            // Determine the row where the new TextBox will be placed (after the last used row).
            int newRow = sheet.Cells.MaxDataRow + 1;

            // Clone the template TextBox using the correct AddCopy overload.
            Shape clonedShape;
            try
            {
                clonedShape = sheet.Shapes.AddCopy(
                    templateTextBox,
                    templateTextBox.UpperLeftRow,
                    templateTextBox.UpperLeftColumn,
                    templateTextBox.Height,
                    templateTextBox.Width);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cloning TextBox: {ex.Message}");
                return;
            }

            // Cast the cloned shape back to TextBox.
            TextBox? newTextBox = clonedShape as TextBox;
            if (newTextBox == null)
            {
                Console.WriteLine("Failed to cast cloned shape to TextBox.");
                return;
            }

            // Update the placeholder text of the cloned TextBox.
            newTextBox.Text = "Your new placeholder text";

            // Position the new TextBox in the desired row while keeping the original column.
            newTextBox.UpperLeftRow = newRow;
            newTextBox.UpperLeftColumn = templateTextBox.UpperLeftColumn;

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
