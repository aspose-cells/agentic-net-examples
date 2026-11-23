// Title: Find the worksheet that contains a specific named shape in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find worksheet that holds a shape with a given name | search all worksheets for a shape named "MyTargetShape" using Aspose.Cells .NET | iterate through workbook shapes to locate a specific shape in C# | how to get the worksheet name of a shape in Aspose.Cells | case‑insensitive shape name lookup across worksheets Aspose.Cells
// Tags: iterate worksheets Aspose.Cells shape lookup | find shape by name Aspose.Cells .NET | retrieve worksheet containing specific shape | case‑insensitive shape name comparison C# | Aspose.Cells workbook shape search

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example verifies the Excel file exists, loads it with Aspose.Cells, then loops through each worksheet and its Shapes collection to locate a shape whose Name matches the target (ignoring case). When found, it records the containing worksheet and outputs its name; otherwise it reports that the shape was not found.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Name of the shape we are looking for
            string targetShapeName = "MyTargetShape";

            // Variable to hold the worksheet that contains the target shape
            Worksheet targetWorksheet = null;

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Check if the shape's name matches the target name
                    if (shape.Name.Equals(targetShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetWorksheet = ws; // Found the worksheet
                        break; // Exit inner loop
                    }
                }

                if (targetWorksheet != null)
                    break; // Exit outer loop once the worksheet is found
            }

            // Output the result
            if (targetWorksheet != null)
            {
                Console.WriteLine($"Shape \"{targetShapeName}\" found in worksheet: {targetWorksheet.Name}");
            }
            else
            {
                Console.WriteLine($"Shape \"{targetShapeName}\" not found in any worksheet.");
            }

            // (Optional) Save the workbook if any modifications were made (lifecycle rule: save)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
