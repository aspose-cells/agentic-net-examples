// Title: Delete all pivot tables from an Excel worksheet and verify their removal with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove all pivot tables from a worksheet | How to check if a worksheet still contains pivot tables after deletion using Aspose.Cells | Aspose.Cells .NET verify pivot table collection is empty after removal | Iterate and delete pivot tables in an Excel file with Aspose.Cells C# | Sample code to delete pivot tables and confirm none remain in Aspose.Cells
// Tags: aspocells delete pivot tables | excel worksheet pivot table removal c# | aspocells verify empty pivot collection | c# loop remove pivot tables aspocells | save workbook after pivot deletion aspocells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one), ensures a worksheet is present, iteratively removes every pivot table from the first worksheet using a while loop, checks that the PivotTables collection is empty, outputs the verification result, and saves the updated workbook to a new file.
class PivotTableDeletionCheck
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    // If loading fails (e.g., password protected), create a new workbook.
                    Console.WriteLine($"Failed to load \"{inputPath}\": {ex.Message}");
                    Console.WriteLine("Creating a new workbook instead.");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Remove all pivot tables from the worksheet.
            while (sheet.PivotTables.Count > 0)
            {
                sheet.PivotTables.RemoveAt(0);
            }

            // Verify removal.
            bool hasPivotTables = sheet.PivotTables.Count > 0;
            Console.WriteLine("Pivot tables present after deletion: " + hasPivotTables);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
