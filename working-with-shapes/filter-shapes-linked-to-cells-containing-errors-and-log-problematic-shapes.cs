// Title: C# example: Filter Excel shapes linked to error cells and log details with Aspose.Cells
// Common Searches: Aspose.Cells C# filter shapes that reference cells with #N/A error | How to log Excel shape information when the linked cell contains an error using Aspose.Cells | Iterate worksheet shapes and detect linked cell error values in .NET | Retrieve and report shapes linked to error cells in an Excel workbook with Aspose.Cells | C# code to find shapes with LinkedCell pointing to error cells in Excel
// Tags: shape linked cell error detection Aspose.Cells | log problematic shapes in Excel workbook C# | worksheet shape enumeration Aspose.Cells .NET | cell error type extraction via shape link | Aspose.Cells shape error handling workflow

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook (or creates a new one), walks through each worksheet and its collection of shapes, checks whether a shape has a LinkedCell, retrieves that cell, and if the cell's value type is an error (CellValueType.IsError) it writes the worksheet name, shape name, linked cell address, and the error value to the console. Shape‑processing exceptions are caught and reported, and the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Iterate through each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape in the worksheet.
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process shapes that are linked to a cell (LinkedCell is non‑empty).
                        if (!string.IsNullOrEmpty(shape.LinkedCell))
                        {
                            string linkedCellName = shape.LinkedCell;

                            // Retrieve the linked cell.
                            Cell linkedCell = sheet.Cells[linkedCellName];

                            // Check if the cell contains an error value.
                            if (linkedCell.Type == CellValueType.IsError)
                            {
                                Console.WriteLine(
                                    $"Worksheet: {sheet.Name}, Shape: {shape.Name}, Linked Cell: {linkedCellName}, Error Value: {linkedCell.StringValue}");
                            }
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape \"{shape.Name}\": {exShape.Message}");
                    }
                }
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
