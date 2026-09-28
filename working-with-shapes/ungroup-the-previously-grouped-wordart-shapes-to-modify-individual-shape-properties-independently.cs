// Title: How to ungroup WordArt GroupShape objects in an Excel workbook and edit each shape’s fill color, text, and rotation using Aspose.Cells for .NET
// AI Prompts: Load an Excel file with Aspose.Cells, locate all GroupShape objects, call Ungroup on each, then iterate the resulting Shape collection to set FillFormat.ForeColor, update the Text property, and assign a RotationAngle before saving. | Using C#, ungroup WordArt shapes in a worksheet, change each shape’s background color to LightBlue, replace its text with "Updated Text", and rotate it 15 degrees with Aspose.Cells.
// Common Searches: Aspose.Cells C# ungroup WordArt shapes and change fill color | How to modify individual shape properties after ungrouping in Excel with .NET | Set rotation angle for Excel shapes using Aspose.Cells | Iterate through worksheet Shapes collection after Ungroup in C# | Update text of WordArt shapes in an Excel file with Aspose.Cells
// Tags: ungroup GroupShape Aspose.Cells .NET | modify shape fill color Aspose.Cells | change WordArt text property C# | apply rotation angle to Excel shape | iterate worksheet shapes after ungroup

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, scans the first worksheet for GroupShape objects, calls Ungroup to break them into separate shapes, then loops through all shapes to set a light blue fill color, change the text to "Updated Text", apply a 15‑degree rotation, and finally saves the modified workbook.
class UngroupWordArtExample
{
    static void Main()
    {
        try
        {
            const string inputFile = "GroupedWordArt.xlsx";
            const string outputFile = "UngroupedWordArt.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook containing the grouped WordArt shapes
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Scan for GroupShape objects and ungroup them
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                if (shape is GroupShape groupShape)
                {
                    try
                    {
                        // Ungroup the shapes; inner shapes are added to the worksheet's Shapes collection
                        groupShape.Ungroup();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to ungroup shape at index {i}: {ex.Message}");
                    }

                    // Restart scanning because the collection has changed
                    i = -1;
                }
            }

            // Modify each individual shape as needed
            foreach (Shape shape in sheet.Shapes)
            {
                // Set fill color if the shape supports FillFormat
                try
                {
                    if (shape.FillFormat != null)
                    {
                        shape.FillFormat.ForeColor = Color.LightBlue;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set fill color for shape: {ex.Message}");
                }

                // Update text for shapes that have a Text property
                try
                {
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        shape.Text = "Updated Text";
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set text for shape: {ex.Message}");
                }

                // Set rotation angle (available on the base Shape class)
                try
                {
                    shape.RotationAngle = 15;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set rotation for shape: {ex.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as '{outputFile}'.");
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
