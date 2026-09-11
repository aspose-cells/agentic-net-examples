// Title: Remove all CheckBox form controls linked to cells with value zero from every worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loops through each worksheet, finds CheckBox shapes whose LinkedCell contains the numeric value 0, and deletes those shapes. | Show how to collect CheckBox controls whose linked cells evaluate to zero into a list and remove them from the ShapeCollection in an Excel workbook using Aspose.Cells. | Provide a method that safely checks for a linked cell, parses its value, and eliminates the corresponding CheckBox shape across all sheets in a .xlsx file.
// Common Searches: aspnet aspose.cells delete checkbox form control when linked cell equals 0 | c# iterate all worksheets and remove checkboxes linked to zero values | how to programmatically remove Excel checkboxes based on linked cell content using Aspose.Cells | remove form control checkboxes with zero linked cell in .xlsx via Aspose.Cells .NET
// Tags: delete checkbox shapes linked cell zero Aspose.Cells | iterate worksheets shape collection .NET | form control removal based on cell value | Aspose.Cells conditional shape deletion | C# Excel checkbox linked cell processing

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, iterates through each worksheet, identifies CheckBox form controls whose LinkedCell contains the numeric value 0, removes those shapes from the worksheet's ShapeCollection, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;
                List<Shape> shapesToDelete = new List<Shape>();

                // Examine each shape
                foreach (Shape shape in shapes)
                {
                    // Process only CheckBox form controls
                    if (shape is CheckBox checkBox)
                    {
                        string linkedCellAddress = checkBox.LinkedCell;
                        if (string.IsNullOrEmpty(linkedCellAddress))
                            continue;

                        // Access the linked cell
                        Cell linkedCell = sheet.Cells[linkedCellAddress];
                        if (linkedCell?.Value == null)
                            continue;

                        // Check if the linked cell contains numeric zero
                        if (double.TryParse(linkedCell.Value.ToString(), out double cellValue) && cellValue == 0)
                        {
                            shapesToDelete.Add(shape);
                        }
                    }
                }

                // Remove identified CheckBox controls
                foreach (Shape shape in shapesToDelete)
                {
                    shapes.Remove(shape);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
