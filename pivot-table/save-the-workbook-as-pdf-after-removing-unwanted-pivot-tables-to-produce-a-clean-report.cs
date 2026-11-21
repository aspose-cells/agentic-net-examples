// Title: Remove specific pivot tables from an Excel workbook and save as PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete pivot table by name before PDF export | How to programmatically remove selected pivot tables from an .xlsx file using Aspose.Cells | Export cleaned Excel workbook to PDF after removing unwanted pivot tables in C# | C# code to filter out specific pivot tables in a workbook with Aspose.Cells | Remove multiple pivot tables from worksheets and save as PDF using Aspose.Cells .NET
// Tags: remove specific pivot tables Aspose.Cells | pivot table deletion C# Aspose.Cells | save workbook as PDF Aspose.Cells | clean Excel report PDF export | filter pivot tables by name Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // PivotTable related classes

// The example loads an existing Excel file, defines a list of pivot table names to discard, iterates through each worksheet to remove matching PivotTable objects, and then saves the cleaned workbook as a PDF, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Names of pivot tables that should be removed
            HashSet<string> pivotTablesToRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "PivotTable1",
                "PivotTable2"
                // add more names as needed
            };

            // Iterate through each worksheet and remove matching pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PivotTableCollection pivots = sheet.PivotTables;

                // Remove in reverse order to prevent index shifting
                for (int i = pivots.Count - 1; i >= 0; i--)
                {
                    PivotTable pt = pivots[i];
                    if (pivotTablesToRemove.Contains(pt.Name))
                    {
                        pivots.RemoveAt(i);
                    }
                }
            }

            // Save the cleaned workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
