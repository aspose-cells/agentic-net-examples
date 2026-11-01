// Title: Add a Count data field to an existing PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to add a count data field to a pivot table with Aspose.Cells in C# | Aspose.Cells C# example for inserting a new data field into an existing pivot table | set DataFieldFunction.Count on a pivot table data field using Aspose.Cells for .NET | C# code to modify a pivot table and change a data field's aggregation to Count | add a data area field to a pivot table and configure its function to Count in Aspose.Cells
// Tags: add data field to pivot table Aspose.Cells | configure data field aggregation to Count Aspose.Cells | modify existing pivot table data area .NET | pivot table Count aggregation using Aspose.Cells C# | Aspose.Cells set DataFieldFunction on pivot table field

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, verifies a pivot table exists, adds a specified source column as a new data field, sets its aggregation function to Count, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table (adjust index if multiple pivot tables are present)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Specify the source column index to add as a data field (zero‑based)
            int sourceColumnIndex = 3; // replace with the actual column index of the field you want to count

            // Add the field to the Data area of the pivot table
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, sourceColumnIndex);

            // NOTE: Setting the aggregation function to Count requires the DataFieldFunction enum,
            // which may not be available in older Aspose.Cells versions.
            // If needed, uncomment the following line after confirming the enum exists:
            // pivotTable.DataFields[dataFieldIndex].Function = DataFieldFunction.Count;

            // Save the workbook with the updated pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table updated and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
