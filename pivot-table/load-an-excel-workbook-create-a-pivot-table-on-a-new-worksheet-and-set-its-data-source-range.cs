// Title: C# – Add a Pivot Table to a New Worksheet from an Existing or Generated Workbook with Aspose.Cells
// Description: Loads an Excel file (or creates a simple workbook if missing), names the first sheet "SourceData", determines its used range, adds a "PivotTable" sheet, creates a pivot table named "MyPivotTable" using that range, maps the first column to rows and the second to data, and saves the result as Output.xlsx.
// Keywords: Aspose.Cells pivot table C# | create pivot table .NET | dynamic source range Aspose.Cells | add worksheet pivot table | load or generate workbook Aspose | MaxDisplayRange pivot source | C# Excel automation Aspose
// Common Searches: Aspose.Cells add pivot table C# example | Create pivot table from used range Aspose.Cells | C# generate workbook if file not found Aspose | Set pivot table source data programmatically | Pivot table on new sheet Aspose.Cells .NET
// Developer Intent: Generate a pivot table on a newly added worksheet, using a dynamically calculated data range from an existing or auto‑created workbook.
// Use Cases: Automatically summarize imported data by inserting a pivot table even when the source file is absent. | Build monthly sales or inventory reports that create a pivot table on a separate sheet with row and value fields mapped automatically. | Provide a fallback workbook with sample data for testing pipelines that require a pivot table output.
// AI Prompts: Show how to add multiple data fields to the pivot table created in this example. | Explain how to apply currency formatting to the pivot table values using Aspose.Cells. | Give code to refresh the pivot table after programmatically updating the source data.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel file (or creates a simple workbook if missing), names the first sheet "SourceData", determines its used range, adds a "PivotTable" sheet, creates a pivot table named "MyPivotTable" using that range, maps the first column to rows and the second to data, and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "SourceData";
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
            }

            // Ensure the source worksheet is named "SourceData"
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Determine the used range of the source data (fully qualified to avoid ambiguity)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"={sourceSheet.Name}!{sourceRange.Address}";

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "MyPivotTable");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields: first column as row, second as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Save the workbook with the new pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table created successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
