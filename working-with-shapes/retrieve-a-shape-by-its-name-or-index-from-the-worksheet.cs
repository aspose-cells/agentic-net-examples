// Title: Retrieve a worksheet shape by name or by index using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get shape by name from worksheet | C# retrieve Excel shape using index with Aspose.Cells | Iterate worksheet.Shapes collection to find a specific shape in Aspose.Cells | Find shape position in Excel file using Aspose.Cells .NET API | Check shape index bounds before accessing in Aspose.Cells C#
// Tags: worksheet shape lookup by name Aspose.Cells | excel shape index access Aspose.Cells C# | shape collection iteration Aspose.Cells | c# retrieve excel shape without GetByName | aspnet excel shape handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, accesses the first worksheet, and demonstrates two ways to obtain a shape: by iterating the Shapes collection to match a given name, and by directly indexing the collection after validating the index range. It prints the found shape's name and type without modifying or saving the file.
class RetrieveShapeExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ----- Retrieve shape by its name -----
            string shapeName = "MyShape";
            Shape shapeByName = null;

            try
            {
                // ShapeCollection does not provide GetByName; iterate manually
                foreach (Shape shp in worksheet.Shapes)
                {
                    if (shp.Name.Equals(shapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        shapeByName = shp;
                        break;
                    }
                }

                if (shapeByName != null)
                {
                    Console.WriteLine($"Shape found by name: {shapeByName.Name}, Type: {shapeByName.Type}");
                }
                else
                {
                    Console.WriteLine($"No shape with name '{shapeName}' was found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving shape by name: {ex.Message}");
            }

            // ----- Retrieve shape by its index -----
            int shapeIndex = 0; // example index
            Shape shapeByIndex = null;

            if (shapeIndex >= 0 && shapeIndex < worksheet.Shapes.Count)
            {
                try
                {
                    shapeByIndex = worksheet.Shapes[shapeIndex];
                    Console.WriteLine($"Shape found at index {shapeIndex}: {shapeByIndex.Name}, Type: {shapeByIndex.Type}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving shape by index: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Shape index {shapeIndex} is out of range. Total shapes: {worksheet.Shapes.Count}");
            }

            // No modifications are made; saving is optional.
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
