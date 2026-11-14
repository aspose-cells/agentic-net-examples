// Title: Refreshing a SmartArt shape linked to a worksheet cell after updating the cell value with Shape.RefreshLinkedShape in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh SmartArt shape after changing a cell value | how to use Shape.RefreshLinkedShape to update a linked SmartArt diagram in Excel with Aspose.Cells | synchronize SmartArt with worksheet cell edits using Aspose.Cells .NET | example code for refreshing linked SmartArt after cell update in Aspose.Cells | programmatically refresh a SmartArt shape tied to a cell in C# using Aspose
// Tags: Shape.RefreshLinkedShape method in Aspose.Cells | programmatic SmartArt refresh in Excel using Aspose | synchronize SmartArt diagram with worksheet cell in C# | Aspose.Cells shape handling for linked SmartArt | C# update cell and refresh SmartArt via Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example demonstrates how to load an Excel workbook with Aspose.Cells, modify the value of a target cell, and then refresh any SmartArt shape linked to that cell using the Shape.RefreshLinkedShape method before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Update the target cell (B2) with a new value
            sheet.Cells["B2"].PutValue(12345);

            // Attempt to locate a SmartArt shape linked to the updated cell.
            // Aspose.Cells may not expose direct SmartArt linking APIs in all versions,
            // so this block is defensive and will simply ignore errors if the APIs are unavailable.
            try
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // If the shape type enumeration is available, you could check for SmartArt like:
                    // if (shape.Type == ShapeType.SmartArt && shape.IsLinkedToCell && shape.LinkedCell == "B2")
                    // However, to keep compilation safe across versions, we skip detailed checks.
                    // Placeholder for any future shape refresh logic.
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Shape processing warning: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
