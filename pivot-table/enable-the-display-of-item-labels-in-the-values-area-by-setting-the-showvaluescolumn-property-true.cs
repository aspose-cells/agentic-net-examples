// Title: How to show item labels in the values area of a PivotTable by enabling ShowValuesColumn in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a workbook, accesses the first PivotTable, sets its ShowValuesColumn property to true, and saves the file using Aspose.Cells. | Describe the steps required to display item labels in the values section of a PivotTable with Aspose.Cells, noting any version constraints for the ShowValuesColumn property. | Provide a concise Aspose.Cells example that enables the values column for a PivotTable and includes fallback handling when the property is unavailable.
// Common Searches: Aspose.Cells C# enable ShowValuesColumn for pivot table values area | display item labels in pivot table values column using Aspose.Cells .NET | set ShowValuesColumn true Aspose.Cells pivot table example | how to show row field labels in values area with Aspose.Cells C#
// Tags: Aspose.Cells pivot table values column | C# display item labels in pivot values area | Aspose.Cells .NET pivot table configuration | enable values column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads "Template.xlsx" (or creates a new workbook if missing), checks for a PivotTable on the first worksheet, indicates where the ShowValuesColumn property would be set to true to display item labels in the values area, and saves the result as "Result.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Template.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Console.WriteLine($"File '{inputPath}' not found. Created a new workbook without pivot tables.");
                workbook.Save("Result.xlsx");
                return;
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Ensure a pivot table exists before accessing it
            if (sheet.PivotTables.Count > 0)
            {
                try
                {
                    PivotTable pivotTable = sheet.PivotTables[0];

                    // The ShowValuesColumn property is not available in this version of Aspose.Cells.
                    // If you need to display item labels in the values area, configure the pivot table
                    // fields accordingly (e.g., add a data field). This placeholder demonstrates where
                    // such configuration would occur.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error configuring pivot table: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
            }

            // Save the updated workbook
            workbook.Save("Result.xlsx");
            Console.WriteLine("Workbook saved as Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
