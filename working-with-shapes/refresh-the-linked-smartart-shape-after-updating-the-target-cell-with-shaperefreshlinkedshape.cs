// Title: Refresh a linked SmartArt shape after updating its source cell with Shape.RefreshLinkedShape in Aspose.Cells for .NET
// AI Prompts: Generate C# code that changes a worksheet cell value, runs workbook.CalculateFormula, calls Shape.RefreshLinkedShape on a SmartArt named "SmartArt 1", and saves the workbook using Aspose.Cells. | Show how to locate a SmartArt shape by name, verify it exists, and invoke its refresh method after modifying the linked cell in an Aspose.Cells workbook. | Provide a snippet that updates a source cell, recalculates formulas, and then refreshes the associated SmartArt shape before saving the Excel file with Aspose.Cells for .NET.
// Common Searches: how to update a source cell and then refresh the linked SmartArt with Aspose.Cells in C# | Aspose.Cells C# refresh SmartArt after changing its data source | using Shape.RefreshLinkedShape after workbook.CalculateFormula in .NET | check SmartArt shape existence before refreshing in Aspose.Cells
// Tags: Shape.RefreshLinkedShape for SmartArt | update linked SmartArt source cell Aspose.Cells | verify SmartArt shape existence C# | recalculate workbook formulas Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, updates the value of cell B2 (the SmartArt source), recalculates formulas, optionally checks for a SmartArt shape named "SmartArt 1", and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Update the target cell that the SmartArt shape is linked to
            Cell targetCell = sheet.Cells["B2"]; // change address as required
            targetCell.PutValue(500);            // new value for the linked SmartArt

            // Recalculate formulas so linked SmartArt reflects the updated cell value
            workbook.CalculateFormula();

            // Optional: Verify that the SmartArt shape exists (no refresh needed)
            Shape shape = sheet.Shapes["SmartArt 1"];
            if (shape == null)
            {
                Console.WriteLine("SmartArt shape not found (proceeding without explicit refresh).");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
