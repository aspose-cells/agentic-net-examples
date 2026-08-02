// Title: C# – Add a Pivot Table to a New Worksheet from an Existing Excel File using Aspose.Cells
// Description: Loads an existing workbook, gets the used range of the first sheet, creates a new worksheet, inserts a pivot table at A1 with that range, assigns the first column as a row field and the second as a data field, and saves the updated file.
// Keywords: Aspose.Cells C# pivot table | create pivot table .NET | pivot table source range Aspose | add pivot table to new sheet | load workbook Aspose.Cells | Excel pivot table automation C#
// Common Searches: how to create a pivot table with Aspose.Cells C# | set pivot table source range using MaxDisplayRange | add pivot table to a new worksheet in .NET | Aspose.Cells example for pivot table creation | C# code to generate pivot table from existing workbook
// Developer Intent: Insert a pivot table on a newly added worksheet based on the used range of an existing sheet and save the workbook.
// Use Cases: Generate a sales summary by adding a pivot table to a template workbook. | Automate monthly reporting: inject a pivot table into uploaded Excel files before archiving. | Create dynamic dashboards that add pivot tables to user‑provided spreadsheets on the fly.
// AI Prompts: Show how to add multiple row or column fields to the pivot table created with Aspose.Cells. | Provide code to refresh the pivot table after changing the source data and before saving. | Explain how to apply a built‑in style and format data fields of the pivot table using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing workbook, gets the used range of the first sheet, creates a new worksheet, inserts a pivot table at A1 with that range, assigns the first column as a row field and the second as a data field, and saves the updated file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (first sheet) containing the data
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Determine the used range of the source data (use Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;

            // Build the source data string in the required format
            string sourceData = $"={sourceSheet.Name}!{sourceRange.Address}";

            // Add a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

            // Configure the pivot table fields (first column as row, second as data)
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Save the workbook with the new pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table created and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
