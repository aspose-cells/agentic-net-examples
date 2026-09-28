// Title: Log each pivot group identifier after grouping a numeric field with Aspose.Cells for .NET
// AI Prompts: Write C# code that groups a numeric PivotField by a given interval, then enumerates the resulting PivotItems and writes each group label to the console for audit logging. | Show how to persist an Aspose.Cells workbook while capturing and outputting pivot group IDs after a GroupBy operation, including directory creation and exception handling.
// Common Searches: Aspose.Cells C# how to get pivot group names after using GroupBy | C# log pivot table groups for audit with Aspose.Cells | retrieve PivotItem.Name after grouping numeric field in Aspose.Cells | save workbook and output pivot group identifiers using Aspose.Cells .NET | audit callback after each pivot group processed in Aspose.Cells
// Tags: Aspose.Cells pivot GroupBy logging | C# PivotItems enumeration after GroupBy | audit pivot groups Aspose.Cells | record group identifiers .NET Excel | Aspose.Cells workbook persistence post grouping

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGroupCallbackDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, groups a numeric row field by an interval, iterates through the generated PivotItems to output each group label for auditing, and saves the workbook to an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (numeric values to be grouped)
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Value";

                sheet.Cells["A2"].Value = "Item1";
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["A3"].Value = "Item2";
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["A4"].Value = "Item3";
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["A5"].Value = "Item4";
                sheet.Cells["B5"].Value = 4;
                sheet.Cells["A6"].Value = "Item5";
                sheet.Cells["B6"].Value = 5;
                sheet.Cells["A7"].Value = "Item6";
                sheet.Cells["B7"].Value = 6;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "DemoPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Add the numeric field as a row field (to be grouped)
                pivot.AddFieldToArea(PivotFieldType.Row, "Value");

                // Add the category field as a data field (just for completeness)
                pivot.AddFieldToArea(PivotFieldType.Data, "Category");

                // Retrieve the row field that will be grouped
                PivotField valueField = pivot.RowFields[0];

                // Group the numeric field by an interval of 2 (creates groups: 1-2, 3-4, 5-6)
                // The second parameter 'false' indicates that the grouping will be applied to the existing field
                valueField.GroupBy(2.0, false);

                // Recalculate the pivot table after grouping
                pivot.CalculateData();

                // Callback simulation: after grouping, log each group identifier
                Console.WriteLine("Audit Log - Grouping Results:");
                foreach (PivotItem item in valueField.PivotItems)
                {
                    // Each PivotItem.Name represents the group label (e.g., "1-2")
                    Console.WriteLine($"Group ID: {item.Name}");
                }

                // Determine output path and ensure directory exists
                string outputPath = "GroupCallbackAuditDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".";
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
