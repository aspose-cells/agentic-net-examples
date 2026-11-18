// Title: How to unlock all shapes in an Excel worksheet using Aspose.Cells for .NET by iterating the ShapeCollection and clearing the IsLocked property
// Common Searches: Aspose.Cells C# unlock all shapes in a worksheet | set shape.IsLocked = false for each shape in Aspose.Cells | remove shape protection from an Excel file using Aspose.Cells .NET | iterate worksheet.Shapes collection to change locking in C# | unlock drawing objects in Excel with Aspose.Cells API
// Tags: unlock worksheet shapes Aspose.Cells | iterate ShapeCollection C# Aspose.Cells | clear shape IsLocked property .NET | modify Excel shape locking Aspose.Cells | Aspose.Cells shape property API

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet, loops through its Shapes collection, sets each shape's IsLocked property to false, and saves the workbook to a new file, demonstrating how to programmatically unlock all drawing objects using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet and unlock them
            foreach (Shape shape in worksheet.Shapes)
            {
                // Correct property to control locking of a shape
                shape.IsLocked = false;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
