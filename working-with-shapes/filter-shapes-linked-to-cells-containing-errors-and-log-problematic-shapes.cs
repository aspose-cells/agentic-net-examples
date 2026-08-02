// Title: Aspose.Cells .NET – Find and Log Shapes Linked to Error Cells
// Description: Loads an Excel workbook (optionally ignoring useless shapes), iterates each worksheet’s ShapeCollection, uses Shape.GetLinkedCell to obtain the linked cell address, checks Cell.IsErrorValue, writes the shape name and type for error‑linked shapes to the console, and saves the workbook unchanged.
// Keywords: Aspose.Cells .NET shape error detection | filter shapes by linked cell error | GetLinkedCell error value | log shapes linked to error cells | ignore useless shapes Aspose.Cells | C# Excel shape validation | Aspose.Cells workbook processing | detect #VALUE! #N/A linked shapes
// Common Searches: Aspose.Cells find shapes linked to error cells | C# iterate Excel shapes and check linked cell for errors | log shape name when linked cell contains #VALUE! | ignore useless shapes while loading workbook Aspose.Cells | how to detect error values in linked cells using Aspose.Cells
// Developer Intent: Locate any shape whose linked cell contains an error value and output its details for review or further processing.
// Use Cases: Create a pre‑publish audit that lists all charts, pictures, or buttons pointing to error cells. | Validate a spreadsheet template by flagging shapes linked to erroneous formulas before distribution. | Automate cleanup of workbooks by identifying and optionally removing shapes associated with error cells.
// AI Prompts: Generate a method that returns a collection of shape names and types linked to error cells using Aspose.Cells. | Extend the example to also capture shapes linked to #N/A cells and export the results to a CSV file. | Add comprehensive error handling for cases where GetLinkedCell returns null, an empty string, or an invalid address.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook (optionally ignoring useless shapes), iterates each worksheet’s ShapeCollection, uses Shape.GetLinkedCell to obtain the linked cell address, checks Cell.IsErrorValue, writes the shape name and type for error‑linked shapes to the console, and saves the workbook unchanged.
    public class FilterErrorLinkedShapes
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook with options to ignore useless shapes (optional but demonstrates LoadOptions usage)
                LoadOptions loadOptions = new LoadOptions
                {
                    IgnoreUselessShapes = true
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = worksheet.Shapes;

                    // Examine each shape in the worksheet
                    foreach (Shape shape in shapes)
                    {
                        // Retrieve the linked cell address (if any) for the shape
                        string linkedCellAddress = shape.GetLinkedCell(true, true);

                        // If the shape is linked to a cell, check whether that cell contains an error
                        if (!string.IsNullOrEmpty(linkedCellAddress))
                        {
                            Cell linkedCell = worksheet.Cells[linkedCellAddress];

                            // Log shapes whose linked cell has an error value
                            if (linkedCell.IsErrorValue)
                            {
                                Console.WriteLine($"Shape '{shape.Name}' (Type={shape.Type}) is linked to error cell '{linkedCellAddress}'.");
                            }
                        }
                    }
                }

                // Save the workbook (unchanged) to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            FilterErrorLinkedShapes.Run();
        }
    }
}
