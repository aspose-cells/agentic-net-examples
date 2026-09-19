// Title: Recursively ungroup nested GroupShape objects on all worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and recursively ungroups every GroupShape until only individual shapes remain. | Create a method for Aspose.Cells that safely scans a worksheet's Shapes collection, detects GroupShape instances, calls Ungroup(), and repeats the scan to handle deeper nesting.
// Common Searches: how to recursively ungroup GroupShape objects in an Excel file using Aspose.Cells C# | flatten nested shape groups across all worksheets with Aspose.Cells .NET | Aspose.Cells C# ungroup all shapes in workbook without losing inner shapes | remove shape groups from Excel workbook programmatically using Aspose.Cells | iterate worksheet shapes and ungroup groups in C# Aspose.Cells example
// Tags: ungroup GroupShape Aspose.Cells | recursive shape group removal C# | remove Excel shape groups Aspose.Cells | iterate worksheet shapes Aspose.Cells | handle shape collection changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an input Excel file, loops through each worksheet, and repeatedly scans the Shapes collection to find GroupShape objects. Each found group is ungrouped with GroupShape.Ungroup(), and the scan restarts to ensure nested groups are also flattened. After all groups are removed, the workbook is saved to the specified output path.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                UngroupAllShapes(sheet);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Recursively ungroup all nested groups on the given worksheet
    private static void UngroupAllShapes(Worksheet sheet)
    {
        bool groupFound;
        do
        {
            groupFound = false;

            // Snapshot of current shapes to avoid collection modification issues
            Shape[] currentShapes = new Shape[sheet.Shapes.Count];
            sheet.Shapes.CopyTo(currentShapes, 0);

            foreach (Shape shape in currentShapes)
            {
                if (shape is GroupShape groupShape)
                {
                    // Ungroup the shape; inner shapes are automatically added to the worksheet
                    groupShape.Ungroup();

                    // A group was processed; another pass may be needed for nested groups
                    groupFound = true;
                    break; // Restart scanning as the collection has changed
                }
            }
        } while (groupFound);
    }
}
