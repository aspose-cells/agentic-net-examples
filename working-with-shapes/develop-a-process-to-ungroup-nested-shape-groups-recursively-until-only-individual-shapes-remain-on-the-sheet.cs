// Title: How to recursively ungroup nested shape groups in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through a worksheet's Shapes collection and repeatedly calls Ungroup on each GroupShape until no grouped shapes remain. | Extend the sample to loop through every worksheet in a workbook and ungroup nested shape groups on each sheet before saving the workbook.
// Common Searches: Aspose.Cells C# ungroup nested group shapes in Excel worksheet | recursive shape ungrouping in Excel using Aspose.Cells .NET API | remove all shape groups from an Excel file with Aspose.Cells programmatically | how to iterate over Shapes collection and ungroup groups in Aspose.Cells
// Tags: Aspose.Cells recursive shape ungrouping | C# Ungroup method for GroupShape in Excel | process all worksheets shape groups Aspose.Cells | Excel shape collection iteration Aspose.Cells | remove nested group shapes .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUngroupShapes
{
    // The example loads an Excel workbook, accesses each worksheet, and repeatedly scans the Shapes collection. Whenever a GroupShape is found, its Ungroup method is invoked. The loop continues until no grouped shapes remain, ensuring only individual shapes are left, after which the workbook is saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any specific worksheet)
                Worksheet sheet = workbook.Worksheets[0];

                // Recursively ungroup all shape groups until only individual shapes remain
                bool groupFound;
                do
                {
                    groupFound = false;

                    // Iterate over the shapes collection.
                    // The collection may change during ungrouping, so break after each ungroup
                    // and restart the loop to get an updated collection.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        if (shape.IsGroup)
                        {
                            // Cast to GroupShape to access Ungroup method
                            if (shape is GroupShape groupShape)
                            {
                                groupShape.Ungroup();   // Ungroup the current group shape
                                groupFound = true;      // Need another pass after modification
                                break;                  // Restart enumeration with updated collection
                            }
                        }
                    }
                } while (groupFound);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
