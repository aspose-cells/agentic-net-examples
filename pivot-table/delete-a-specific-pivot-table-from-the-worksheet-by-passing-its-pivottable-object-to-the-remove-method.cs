// Title: Delete a specific PivotTable from an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove pivot table by object reference | Delete PivotTable named "PivotTable1" from a worksheet using Aspose.Cells | How to programmatically remove a pivot table from an Excel file in .NET | C# code to delete a specific pivot table in a workbook with Aspose.Cells | Remove pivot table from the first worksheet and save the workbook using Aspose.Cells
// Tags: Aspose.Cells PivotTables.Remove usage | C# delete specific pivot table | Excel workbook pivot table removal .NET | pivot table deletion via object reference | Aspose.Cells delete pivot table from worksheet

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a pivot table named "PivotTable1" on the first worksheet, removes that PivotTable using the Worksheet.PivotTables.Remove method, and saves the updated workbook to a new file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found. Operation aborted.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the pivot table by name; ensure it exists before removal
            PivotTable pivot = sheet.PivotTables["PivotTable1"];
            if (pivot == null)
            {
                Console.WriteLine("Pivot table \"PivotTable1\" not found. No action taken.");
            }
            else
            {
                // Remove the pivot table from the worksheet
                sheet.PivotTables.Remove(pivot);
                Console.WriteLine("Pivot table \"PivotTable1\" removed successfully.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
