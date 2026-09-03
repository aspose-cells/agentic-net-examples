// Title: Classify each worksheet in an Excel file as Data‑Only, Shape‑Only, Mixed, or Empty using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through all worksheets, determines whether a sheet contains any non‑empty cells, any shapes, and returns a label such as DataOnly, ShapeOnly, Mixed, or Empty. | Enhance the classifier to also output the number of populated cells and the count of shapes found on each worksheet.
// Common Searches: how to detect shapes on a worksheet using Aspose.Cells C# | classify Excel worksheets by content type with Aspose.Cells .NET | determine if a sheet has only data or only drawings in Aspose.Cells | C# Aspose.Cells check for non‑empty cells and pictures in each worksheet | Aspose.Cells mixed content worksheet detection example
// Tags: worksheet shape detection Aspose.Cells | detect populated cells Aspose.Cells | worksheet content type detection Aspose.Cells | MaxDataRow MaxDataColumn usage Aspose.Cells | Shapes collection enumeration Aspose.Cells | mixed content worksheet analysis Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, loops through every worksheet, uses MaxDataRow/MaxDataColumn to scan for any non‑empty cells, checks the Shapes collection for drawing objects, classifies each sheet as DataOnly, ShapeOnly, Mixed, or Empty, stores the results in a dictionary, and prints the classification for each worksheet.
class WorksheetClassifier
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Store classification results for each worksheet
        var classifications = new Dictionary<string, string>();

        try
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasData = false;   // Indicates presence of cell data
                bool hasShape = false;  // Indicates presence of any shape (pictures, charts, etc.)

                // ----- Check for cell data -----
                // Use MaxDataRow/MaxDataColumn to limit the scan to the used range
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow && !hasData; row++)
                {
                    for (int col = 0; col <= maxCol && !hasData; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        // Consider a cell non‑empty if it contains a non‑null value that is not whitespace
                        if (cell != null && cell.Value != null &&
                            !(cell.Value is string str && string.IsNullOrWhiteSpace(str)))
                        {
                            hasData = true;
                        }
                    }
                }

                // ----- Check for shapes -----
                // The Shapes collection contains all drawing objects on the sheet
                if (sheet.Shapes.Count > 0)
                {
                    hasShape = true;
                }

                // ----- Determine classification -----
                string classification;
                if (hasData && hasShape)
                    classification = "Mixed";
                else if (hasData)
                    classification = "DataOnly";
                else if (hasShape)
                    classification = "ShapeOnly";
                else
                    classification = "Empty";

                classifications[sheet.Name] = classification;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during classification: {ex.Message}");
            return;
        }

        // Output the classification results
        foreach (var kvp in classifications)
        {
            Console.WriteLine($"Worksheet: {kvp.Key}, Classification: {kvp.Value}");
        }

        // Optional: Save the workbook if modifications were made
        // workbook.Save("output.xlsx");
    }
}
