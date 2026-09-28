// Title: C# Aspose.Cells example to ensure every shape references a non‑empty cell address
// AI Prompts: Write C# code using Aspose.Cells that loops through all worksheets, examines each shape’s UpperLeftRow and UpperLeftColumn, converts them to a cell address, and flags shapes whose linked cell is null or contains only whitespace. | Create a routine that collects validation messages for shapes linked to empty cells, writes them to a text file, and includes the worksheet and shape names in each message. | Add comprehensive try‑catch blocks around workbook loading, shape processing, file writing, and workbook saving while implementing the shape‑to‑cell validation with Aspose.Cells.
// Common Searches: aspocells c# validate shape linked cell not empty | how to check if Excel shape points to a blank cell using Aspose.Cells | iterate over worksheet shapes and verify cell value Aspose.Cells C# | generate report of shapes with empty linked cells in Aspose.Cells | error handling for shape validation in Aspose.Cells workbook
// Tags: Aspose.Cells shape linked cell validation | Aspose.Cells iterate worksheet shapes | Aspose.Cells check empty cell for shape | Aspose.Cells generate shape validation report | Aspose.Cells robust error handling for shape processing | Aspose.Cells save workbook after validation

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The program loads an Excel workbook with Aspose.Cells, iterates each worksheet and its shapes, derives the cell address from a shape's UpperLeftRow and UpperLeftColumn, verifies that the linked cell is not null or whitespace, records any violations in a text report, and finally saves the workbook.
class ShapeCellValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string reportPath = "validation_report.txt";

        // Ensure the input file exists before attempting to load
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook (lifecycle rule: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Collect validation errors
        List<string> errors = new List<string>();

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each shape on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Determine the cell address based on the shape's upper‑left corner
                    int row = shape.UpperLeftRow;
                    int column = shape.UpperLeftColumn;
                    string cellAddress = CellsHelper.CellIndexToName(row, column);

                    // Validate that the address refers to a real cell and that the cell is not empty
                    Cell linkedCell = sheet.Cells[cellAddress];
                    if (linkedCell.Value == null || string.IsNullOrWhiteSpace(linkedCell.StringValue))
                    {
                        errors.Add($"Shape '{shape.Name}' on sheet '{sheet.Name}' is linked to an empty cell '{cellAddress}'.");
                    }
                }
                catch (Exception ex)
                {
                    // Record any unexpected errors while processing a shape
                    errors.Add($"Error processing shape '{shape.Name}' on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Write validation results to a text file
        try
        {
            File.WriteAllLines(reportPath, errors);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }

        // Save the workbook (lifecycle rule: save)
        try
        {
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
