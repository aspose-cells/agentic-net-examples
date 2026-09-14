// Title: Delete a TextBox shape from an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that finds the first TextBox shape in a worksheet and removes it using Aspose.Cells. | Create a method that iterates through all shapes in a workbook and deletes every TextBox shape with Aspose.Cells. | Show how to delete a TextBox by its name or index and then save the workbook using Aspose.Cells in C#.
// Common Searches: aspnet delete textbox shape from excel file using aspose.cells | c# remove specific textbox from worksheet with aspose.cells library | how to programmatically delete all textboxes in an Excel workbook using aspose.cells | asp.net core find and delete textbox shape in existing xlsx with aspose.cells | remove shape by type textbox from worksheet using aspose.cells c# example
// Tags: Aspose.Cells remove TextBox shape | C# worksheet.Shapes.RemoveAt textbox | Aspose.Cells delete shape by type | Excel file shape manipulation Aspose.Cells | C# delete textbox from xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, searches the first worksheet for a TextBox shape, removes it with worksheet.Shapes.RemoveAt, and saves the updated file.
class DeleteTextboxExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the required one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Find the first textbox shape in the worksheet
            int textboxIndex = -1;
            foreach (Shape shape in worksheet.Shapes)
            {
                // Compare shape type using its string representation to avoid enum dependency
                if (shape.Type.ToString() == "TextBox")
                {
                    textboxIndex = worksheet.Shapes.IndexOf(shape);
                    break; // Remove only the first found textbox; remove this break to delete all
                }
            }

            // Delete the textbox if it was found
            if (textboxIndex != -1)
            {
                worksheet.Shapes.RemoveAt(textboxIndex);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
