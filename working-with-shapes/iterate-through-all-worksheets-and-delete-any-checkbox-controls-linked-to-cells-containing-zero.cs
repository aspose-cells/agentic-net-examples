// Title: Remove all CheckBox form controls linked to zero-valued cells from every worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, scans each worksheet for CheckBox shapes, evaluates the LinkedCell value, and deletes the shape when the cell contains numeric zero. | Create a reusable method that collects CheckBox shapes whose LinkedCell points to a zero cell, then removes them from the worksheet's Shapes collection without altering the collection during iteration. | Provide an example that logs the addresses of removed CheckBox controls while processing all worksheets in a workbook using Aspose.Cells.
// Common Searches: asp.net aspose.cells delete checkbox shape linked to zero cell | c# remove form control checkbox when linked cell equals 0 using Aspose.Cells | how to iterate all worksheets and delete specific checkbox controls in Excel with Aspose.Cells | safe removal of shapes collection in Aspose.Cells C# | filter checkbox controls by linked cell value Aspose.Cells example
// Tags: Aspose.Cells remove checkbox shapes | C# delete form controls based on linked cell value | Iterate worksheets to filter shapes Aspose.Cells | Safe shape collection modification Aspose.Cells | CheckBox LinkedCell zero removal

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, iterates through every worksheet, finds CheckBox shapes whose LinkedCell points to a numeric cell equal to 0, collects those shapes, removes them safely from the Shapes collection, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "{InputPath}";
            string outputPath = "{OutputPath}";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Collect shapes to remove to avoid modifying the collection while iterating
                    var shapesToRemove = new List<Shape>();

                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Check if the shape is a CheckBox
                        if (shape is CheckBox checkBox)
                        {
                            // Get the cell address linked to the CheckBox (e.g., "A1")
                            string linkedCellAddress = checkBox.LinkedCell;

                            if (!string.IsNullOrEmpty(linkedCellAddress))
                            {
                                // Retrieve the linked cell
                                Cell linkedCell = worksheet.Cells[linkedCellAddress];

                                // Determine if the cell contains numeric zero
                                if (linkedCell.Type == CellValueType.IsNumeric && linkedCell.IntValue == 0)
                                {
                                    // Mark this CheckBox for removal
                                    shapesToRemove.Add(shape);
                                }
                            }
                        }
                    }

                    // Remove the identified CheckBox shapes
                    foreach (Shape shape in shapesToRemove)
                    {
                        worksheet.Shapes.Remove(shape);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
