// Title: How to add a Year‑over‑Year growth calculated field to an existing Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add YoY growth calculated field to pivot table programmatically | Create year over year percentage field in Excel pivot table with Aspose.Cells .NET | Add custom calculated field to existing pivot table using Aspose.Cells API | C# code to insert a YoY growth % field into a pivot table in an XLSX workbook | Aspose.Cells example for adding calculated field to first worksheet pivot table
// Tags: Aspose.Cells add calculated field pivot table | C# pivot table YoY growth percentage | Excel workbook modify pivot table Aspose.Cells | Calculated field formula sales prev sales | Aspose.Cells .NET pivot table data area field

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an XLSX workbook, checks that the first worksheet contains a pivot table, then adds a calculated field named "YoY Growth %" with the formula "=('Sales' - 'Prev Sales') / 'Prev Sales'" to the data area of that pivot table, and saves the modified file.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Assume the pivot table we want to modify is the first one on the sheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Define a calculated field formula for Year‑over‑Year growth percentage
            // Adjust field names ("Sales", "Prev Sales") to match your data source.
            string yoyFormula = "=('Sales' - 'Prev Sales') / 'Prev Sales'";

            // Add the calculated field named "YoY Growth %"
            // The third parameter (true) indicates that the field should be added to the data area.
            pivotTable.AddCalculatedField("YoY Growth %", yoyFormula, true);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
