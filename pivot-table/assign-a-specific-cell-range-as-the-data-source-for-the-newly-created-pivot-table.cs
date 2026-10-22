// Title: Assign a specific cell range as the data source for a newly created pivot table with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a new worksheet, selects the cells A1 through D100 as the pivot table data area, and creates a pivot table at cell A1 using Aspose.Cells. | Create a program that checks for an existing Excel file, generates a sample workbook if missing, loads it, adds a pivot sheet, assigns the pivot table's data area, and saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# define pivot table data area A1:D100 | how to set the data range for a pivot table in Aspose.Cells .NET | C# example of creating a pivot table from a specific cell range using Aspose.Cells | add pivot table to worksheet with custom data area Aspose.Cells | Aspose.Cells pivot table source range not whole sheet
// Tags: Aspose.Cells pivot table data area C# | define pivot table source range Aspose.Cells | add pivot table to new worksheet .NET | create pivot table from specific cells Excel | set pivot table data range using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example checks for an input.xlsx file, creates a simple workbook with sample data if it doesn't exist, loads the workbook, adds a new worksheet named "Pivot", defines the source range "A1:D100" as the pivot table's data area, creates a pivot table at cell A1 called "PivotTable1", and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple one if it does not.
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var sheet = sampleWb.Worksheets[0];
                sheet.Name = "Data";

                // Headers
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Quantity");
                sheet.Cells["D1"].PutValue("Price");

                // Sample data
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["D2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Food");
                sheet.Cells["B3"].PutValue("Bread");
                sheet.Cells["C3"].PutValue(5);
                sheet.Cells["D3"].PutValue(1.2);

                sampleWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Get the worksheet named "Data".
            var dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
                throw new Exception("Worksheet 'Data' not found.");

            // Add a new worksheet for the pivot table.
            var pivotSheet = workbook.Worksheets.Add("Pivot");

            // Define the source range for the pivot table.
            string sourceRange = "A1:D100";

            // Create the pivot table at cell A1 (row 0, column 0).
            var pivotTable = pivotSheet.PivotTables.Add(sourceRange, 0, 0, "PivotTable1");

            // (Optional) Configure pivot fields here if needed.

            // Save the workbook with the new pivot table.
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table created and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
