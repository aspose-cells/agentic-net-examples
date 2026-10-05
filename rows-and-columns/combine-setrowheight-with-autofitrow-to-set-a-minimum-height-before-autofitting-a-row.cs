// Title: How to set a minimum row height then auto‑fit a row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that assigns a fixed minimum height to a worksheet row, calls AutoFitRows, and ensures the final height does not drop below the minimum with Aspose.Cells. | Generate a complete example that loads an existing Excel file, sets row 3 height to 20 points, auto‑fits the row, re‑applies the minimum if needed, and saves the workbook using Aspose.Cells in .NET. | Provide a snippet demonstrating how to combine Row.Height and Worksheet.AutoFitRows to enforce a lower bound on row height in a C# Aspose.Cells project.
// Common Searches: Aspose.Cells C# set row height minimum before AutoFitRows | prevent row height from shrinking after auto‑fit in Aspose.Cells | how to enforce lower bound on Excel row height using Aspose.Cells .NET | auto‑fit a specific row while keeping a minimum height in C#
// Tags: row height assignment Aspose.Cells | auto‑fit rows with height floor | minimum row height enforcement C# | worksheet row sizing Aspose.Cells .NET | prevent row shrinkage after autofit

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, sets a minimum height of 20 points for the third row, auto‑fits that row, restores the minimum height if the auto‑fit reduces it, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Row index (zero‑based) and minimum height in points
            int rowIndex = 2;          // third row
            double minHeight = 20.0;   // minimum height

            // Access the row and set its height to the minimum value
            Row row = sheet.Cells.Rows[rowIndex];
            row.Height = minHeight;

            // Auto‑fit the row based on its content
            sheet.AutoFitRows(rowIndex, rowIndex);

            // Ensure the row height does not fall below the minimum
            if (row.Height < minHeight)
            {
                row.Height = minHeight;
            }

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
