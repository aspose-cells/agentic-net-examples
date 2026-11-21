// Title: Sync the checked state of multiple CheckBox shapes with their linked Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set linked cell for each CheckBox shape in a worksheet | how to update Excel cell boolean value based on CheckBox shape state with Aspose.Cells | iterate over worksheet shapes and synchronize CheckBox state in Aspose.Cells .NET | initialize empty cells to false when linking CheckBox controls in Aspose.Cells | save workbook after linking CheckBox shapes to column A cells using Aspose.Cells
// Tags: Aspose.Cells set CheckBox linked cell | Aspose.Cells synchronize CheckBox state with cell | Aspose.Cells iterate worksheet shapes C# | Aspose.Cells initialize boolean cell from CheckBox | Aspose.Cells save workbook after shape processing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, loops through all shapes on the first worksheet, identifies CheckBox shapes, assigns each CheckBox's LinkedCell to the cell in column A of its upper‑left row, ensures the linked cell holds a boolean value (initializing it to false when empty), and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only CheckBox shapes
                    if (shape is CheckBox checkBox)
                    {
                        try
                        {
                            // Row index where the CheckBox is anchored (upper‑left corner)
                            int rowIndex = shape.UpperLeftRow;

                            // Address of the linked cell in column A (e.g., "A2")
                            string linkedCellAddress = sheet.Cells[rowIndex, 0].Name;
                            checkBox.LinkedCell = linkedCellAddress;

                            // Synchronize the cell value with the CheckBox state
                            Cell targetCell = sheet.Cells[rowIndex, 0];
                            if (targetCell.Value == null)
                            {
                                // Initialize the cell with the CheckBox's default state (unchecked)
                                targetCell.Value = false;
                            }
                            else
                            {
                                // Ensure the cell contains a boolean value
                                bool cellValue = Convert.ToBoolean(targetCell.Value);
                                // If needed, additional logic can be added here to reflect the cell value elsewhere
                            }
                        }
                        catch (Exception exShape)
                        {
                            Console.WriteLine($"Error processing CheckBox at row {shape.UpperLeftRow}: {exShape.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
