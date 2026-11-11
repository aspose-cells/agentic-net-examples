// Title: Hide a specific PivotField from an Excel pivot table by removing it from row or column fields using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide a pivot field from the report area | Remove a specific row field from an Excel pivot table with Aspose.Cells .NET | How to programmatically delete a column field in a pivot table using Aspose.Cells | Aspose.Cells hide pivot table field by name in C# | Set pivot field visibility to false with Aspose.Cells for .NET
// Tags: Aspose.Cells pivot field removal C# | Aspose.Cells row field deletion .NET | Aspose.Cells column field elimination C# | Excel pivot table field show hide control | Aspose.Cells API for pivot table manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads a workbook, checks for a pivot table on the first worksheet, searches for a field named "Category" in the row and column fields, removes the matching PivotField to hide it from the report area, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists before loading.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains the pivot table.
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook has at least one worksheet.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one pivot table.
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                PivotTable pivotTable = sheet.PivotTables[0];

                // Name of the pivot field you want to hide from the report area.
                string fieldNameToHide = "Category";

                // Attempt to hide the field by removing it from row or column fields.
                bool fieldRemoved = false;

                // Search in row fields.
                for (int i = 0; i < pivotTable.RowFields.Count; i++)
                {
                    if (pivotTable.RowFields[i].Name.Equals(fieldNameToHide, StringComparison.OrdinalIgnoreCase))
                    {
                        pivotTable.RowFields.RemoveAt(i);
                        fieldRemoved = true;
                        break;
                    }
                }

                // If not found in rows, search in column fields.
                if (!fieldRemoved)
                {
                    for (int i = 0; i < pivotTable.ColumnFields.Count; i++)
                    {
                        if (pivotTable.ColumnFields[i].Name.Equals(fieldNameToHide, StringComparison.OrdinalIgnoreCase))
                        {
                            pivotTable.ColumnFields.RemoveAt(i);
                            fieldRemoved = true;
                            break;
                        }
                    }
                }

                if (fieldRemoved)
                {
                    Console.WriteLine($"Field '{fieldNameToHide}' has been hidden (removed) from the pivot table.");
                }
                else
                {
                    Console.WriteLine($"Field '{fieldNameToHide}' not found in the pivot table.");
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to find a pivot field by name in row or column fields.
        private static PivotField FindPivotField(PivotTable pivotTable, string fieldName)
        {
            foreach (PivotField pf in pivotTable.RowFields)
            {
                if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return pf;
            }

            foreach (PivotField pf in pivotTable.ColumnFields)
            {
                if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return pf;
            }

            return null;
        }
    }
}
