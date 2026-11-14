// Title: Update a linked cell and refresh its associated shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update cell value and refresh linked shape in worksheet | How to recalculate formulas to refresh a shape linked to a cell with Aspose.Cells .NET | Refresh Excel shape after programmatically changing a linked cell using Aspose.Cells | C# Aspose.Cells automatically update picture or chart after cell B2 value change | Refresh shape content after modifying a linked cell in an Aspose.Cells workbook
// Tags: update linked cell Aspose.Cells C# | refresh worksheet shape Aspose.Cells | recalculate formulas Aspose.Cells workbook | shape linked to cell Aspose.Cells .NET | save workbook after shape refresh C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, changes the value of cell B2, calls Workbook.CalculateFormula() to ensure any shape linked to that cell reflects the new value, and then saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or specify by name)
                Worksheet sheet = workbook.Worksheets[0];

                // Update the linked cell value (e.g., cell B2)
                sheet.Cells["B2"].PutValue("New Value");

                // If there are shapes on the sheet, optionally refresh them
                if (sheet.Shapes.Count > 0)
                {
                    // Assuming the first shape is the one linked to the cell
                    Shape linkedShape = sheet.Shapes[0];

                    // In Aspose.Cells for .NET there is no Shape.Update() method.
                    // Recalculating formulas ensures any linked content is refreshed.
                    workbook.CalculateFormula();
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
