// Title: Hide a specific PivotField (e.g., Region) in an Excel pivot table using Aspose.Cells for .NET (C#)
// Common Searches: C# code to hide a row field in an Excel pivot table using Aspose.Cells | Aspose.Cells .NET hide pivot field by name programmatically | Remove a specific field from a pivot table in C# with Aspose.Cells | How to delete a page field from an Excel pivot table using Aspose.Cells API | Set PivotField hidden property to true with Aspose.Cells for .NET
// Tags: Aspose.Cells hide pivot field C# | remove specific field from pivot table .NET | pivot table row field removal Aspose.Cells | Excel pivot field hidden property Aspose | C# Aspose.Cells pivot table field manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, locates the first pivot table, searches for a field named "Region" across row, column, page, and data collections, removes the field from its collection to hide it, and saves the updated workbook.
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

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the pivot field you want to hide
            string fieldNameToHide = "Region";

            // Locate the field among row, column, page, and data fields
            PivotField targetField = null;

            // Row fields
            foreach (PivotField pf in pivotTable.RowFields)
            {
                if (pf.Name == fieldNameToHide)
                {
                    targetField = pf;
                    break;
                }
            }

            // Column fields (if not found yet)
            if (targetField == null)
            {
                foreach (PivotField pf in pivotTable.ColumnFields)
                {
                    if (pf.Name == fieldNameToHide)
                    {
                        targetField = pf;
                        break;
                    }
                }
            }

            // Page fields (if not found yet)
            if (targetField == null)
            {
                foreach (PivotField pf in pivotTable.PageFields)
                {
                    if (pf.Name == fieldNameToHide)
                    {
                        targetField = pf;
                        break;
                    }
                }
            }

            // Data fields (if not found yet)
            if (targetField == null)
            {
                foreach (PivotField pf in pivotTable.DataFields)
                {
                    if (pf.Name == fieldNameToHide)
                    {
                        targetField = pf;
                        break;
                    }
                }
            }

            // Hide (remove) the field if it was found
            if (targetField != null)
            {
                try
                {
                    // Attempt to remove from each possible collection; only one will succeed
                    if (pivotTable.RowFields.Contains(targetField))
                        pivotTable.RowFields.Remove(targetField);
                    else if (pivotTable.ColumnFields.Contains(targetField))
                        pivotTable.ColumnFields.Remove(targetField);
                    else if (pivotTable.PageFields.Contains(targetField))
                        pivotTable.PageFields.Remove(targetField);
                    else if (pivotTable.DataFields.Contains(targetField))
                        pivotTable.DataFields.Remove(targetField);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to hide the field \"{fieldNameToHide}\": {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Field \"{fieldNameToHide}\" not found in the pivot table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
