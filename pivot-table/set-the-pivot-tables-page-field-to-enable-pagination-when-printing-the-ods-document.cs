// Title: Add a page field to a pivot table in an ODS workbook to enable pagination when printing using Aspose.Cells for .NET
// AI Prompts: Insert a page field into the first pivot table of an ODS file and configure it to create page breaks during printing with Aspose.Cells for .NET. | Programmatically enable printed pagination for a pivot table in an ODS workbook by assigning a page field and saving the workbook via the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# add page field to pivot table in ODS for printing pagination | how to enable page breaks for pivot tables when printing ODS files using Aspose.Cells | set pivot table page field programmatically in ODS workbook with Aspose.Cells .NET | C# example of pagination for ODS pivot table using Aspose.Cells
// Tags: Aspose.Cells add page field pivot table ODS | pivot table pagination ODS Aspose.Cells | C# set pivot table page field Aspose.Cells | ODS workbook printing pagination Aspose.Cells | modify existing pivot table Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotField types

namespace AsposeCellsExample
{
    // The sample loads an ODS workbook, checks for a pivot table, adds the first row field as a page field when none exists, optionally enables item drill‑down, and saves the file, allowing the pivot table to generate separate pages when printed.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.ods";
            const string outputPath = "output.ods";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing ODS workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Retrieve the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // If no page field exists, add the first row field as a page field
                if (pivotTable.PageFields.Count == 0 && pivotTable.RowFields.Count > 0)
                {
                    // Add the existing row field object directly (expects PivotField)
                    PivotField rowField = pivotTable.RowFields[0];
                    pivotTable.PageFields.Add(rowField);
                }

                // Enable drill-down for the first page field, if present
                if (pivotTable.PageFields.Count > 0)
                {
                    // In some Aspose.Cells versions, PivotField does not expose EnableItemDrillDown.
                    // If the property exists, it can be set here. Otherwise, this step is skipped.
                    PivotField pageField = pivotTable.PageFields[0];
                    // Uncomment the following line if your version supports it:
                    // pageField.EnableItemDrillDown = true;
                }

                // Save the modified workbook back to ODS format
                workbook.Save(outputPath, SaveFormat.Ods);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
