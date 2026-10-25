// Title: Add a Salesperson report filter (Page field) to an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a page field to an existing PivotTable | How to set a report filter for the Salesperson column in an Excel PivotTable with Aspose.Cells | C# code to add a Salesperson filter to a PivotTable in a workbook using Aspose.Cells | Adding a report filter field to a PivotTable programmatically with Aspose.Cells .NET | Example of using PivotFieldType.Page to filter PivotTable data in C#
// Tags: Aspose.Cells PivotTable page field insertion | Salesperson filter addition to Excel PivotTable | using PivotFieldType.Page in C# | modify PivotTable filter programmatically | save workbook after PivotTable update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, locates the first PivotTable on the 'Data' worksheet, adds the 'Salesperson' column as a Page (report filter) field using PivotTable.AddFieldToArea with PivotFieldType.Page, and saves the modified file, including checks for missing files, worksheet, and PivotTable.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "Report.xlsx";
                string outputPath = "Report_With_Salesperson_Filter.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the worksheet that contains the PivotTable
                Worksheet worksheet = workbook.Worksheets["Data"];
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet 'Data' not found.");
                    return;
                }

                // Ensure there is at least one PivotTable
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found in the worksheet.");
                    return;
                }

                // Get the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Add the "Salesperson" field as a Report Filter (Page field)
                try
                {
                    // Use overload that accepts field name directly
                    pivotTable.AddFieldToArea(PivotFieldType.Page, "Salesperson");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add 'Salesperson' as a page field: {ex.Message}");
                    return;
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
