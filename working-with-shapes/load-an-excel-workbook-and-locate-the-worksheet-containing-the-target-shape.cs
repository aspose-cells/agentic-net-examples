// Title: Find the Worksheet Containing a Specific Shape in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file using Aspose.Cells, scans each worksheet’s Shapes collection, matches a shape by its Name, and retrieves the parent Worksheet via Shape.Worksheet. The sample prints the worksheet name or reports that the shape was not found.
// Keywords: Aspose.Cells C# find shape worksheet | locate shape by name Aspose.Cells | Shape.Worksheet property | search Excel shapes .NET | iterate worksheets Aspose.Cells | Excel shape parent sheet | C# Aspose.Cells shape lookup | retrieve worksheet of shape
// Common Searches: how to get the worksheet of a shape using Aspose.Cells | C# locate shape named MyRectangle in Excel workbook | Aspose.Cells find which sheet contains a specific shape | search for a shape by name across all worksheets | retrieve parent worksheet of an Excel shape in .NET
// Developer Intent: Determine the worksheet that contains a shape with a given name in an Excel workbook.
// Use Cases: Validate that a required diagram or button exists on the correct sheet before processing data. | Programmatically move, resize, or delete a shape after locating its parent worksheet. | Generate a report listing each worksheet and the shapes it contains for documentation or auditing.
// AI Prompts: Write C# code with Aspose.Cells that returns the name of the worksheet holding a shape called 'MyRectangle'. | Show an example that logs all worksheets that contain shapes whose names start with 'Chart_'. | Explain how to handle a missing shape by throwing a custom NotFoundException in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel file using Aspose.Cells, scans each worksheet’s Shapes collection, matches a shape by its Name, and retrieves the parent Worksheet via Shape.Worksheet. The sample prints the worksheet name or reports that the shape was not found.
    public class LocateShapeWorksheetDemo
    {
        public static void Run()
        {
            // Path to the existing Excel file
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Name of the shape we want to locate
                string targetShapeName = "MyRectangle";

                // Variable to hold the worksheet that contains the target shape
                Worksheet targetWorksheet = null;

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the current worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape's name matches the target name
                        if (shape.Name == targetShapeName)
                        {
                            // Use the Shape.Worksheet property to get the containing worksheet
                            targetWorksheet = shape.Worksheet;
                            break;
                        }
                    }

                    // Exit outer loop if the shape has been found
                    if (targetWorksheet != null)
                        break;
                }

                // Output the result
                if (targetWorksheet != null)
                    Console.WriteLine($"Shape '{targetShapeName}' is located in worksheet: {targetWorksheet.Name}");
                else
                    Console.WriteLine($"Shape '{targetShapeName}' was not found in any worksheet.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            LocateShapeWorksheetDemo.Run();
        }
    }
}
