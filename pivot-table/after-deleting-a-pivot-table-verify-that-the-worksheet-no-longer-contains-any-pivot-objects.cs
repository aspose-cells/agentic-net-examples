// Title: Delete the first pivot table from an Excel worksheet and verify its removal with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that removes the first PivotTable from a worksheet, checks that the PivotTables collection is empty, and saves the workbook. | Show how to programmatically delete a PivotTable in a .NET workbook, confirm no PivotTables remain, and output the verification result.
// Common Searches: C# Aspose.Cells how to remove a pivot table and ensure it is deleted | verify that a worksheet has no pivot tables after removal using Aspose.Cells | Aspose.Cells delete first PivotTable and check PivotTables.Count | remove pivot tables programmatically in .NET and confirm removal | Aspose.Cells sample code for deleting pivot tables and validating cleanup
// Tags: Aspose.Cells delete pivot table C# | Aspose.Cells verify pivot table count | C# remove first worksheet pivot table | Aspose.Cells check worksheet for remaining pivot tables | Aspose.Cells save workbook after pivot removal

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook), accesses the first worksheet, removes the first pivot table if present, verifies that the worksheet's PivotTables collection is empty, prints the verification result, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Remove the first pivot table if any exist.
            if (sheet.PivotTables.Count > 0)
            {
                sheet.PivotTables.RemoveAt(0);
            }

            // Verify removal.
            bool noPivotTables = sheet.PivotTables.Count == 0;
            Console.WriteLine(noPivotTables
                ? "All pivot tables have been removed."
                : "Pivot tables still exist on the worksheet.");

            // Save the workbook with changes.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
