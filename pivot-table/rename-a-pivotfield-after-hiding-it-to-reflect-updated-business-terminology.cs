// Title: Programmatically hide a PivotField and rename it to a new business term using Aspose.Cells for .NET (C#)
// Common Searches: C# hide a pivot field and then rename it with Aspose.Cells | How to remove a row or column field from a pivot table and change its name using Aspose.Cells .NET | Aspose.Cells example for hiding a pivot field before updating its name | Rename a pivot table field after hiding it in an Excel workbook with Aspose.Cells C# | Aspose.Cells code to hide a pivot field and set a new field name
// Tags: hide pivot field Aspose.Cells C# | rename pivot field Aspose.Cells | pivot table field removal Aspose.Cells | update pivot field terminology Aspose.Cells | modify pivot table fields Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing workbook, finds the "Region" pivot field in the first pivot table, removes it from its row or column collection to hide it, changes its name to "Territory", and saves the workbook to a new file.
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

            // Get the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Locate the pivot field named "Region" (search row and column fields)
            PivotField pivotField = null;

            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name == "Region")
                {
                    pivotField = field;
                    break;
                }
            }

            if (pivotField == null)
            {
                foreach (PivotField field in pivotTable.ColumnFields)
                {
                    if (field.Name == "Region")
                    {
                        pivotField = field;
                        break;
                    }
                }
            }

            if (pivotField != null)
            {
                // Hide the field by removing it from its collection (row or column)
                if (pivotTable.RowFields.Contains(pivotField))
                {
                    pivotTable.RowFields.Remove(pivotField);
                }
                else if (pivotTable.ColumnFields.Contains(pivotField))
                {
                    pivotTable.ColumnFields.Remove(pivotField);
                }

                // Rename the field
                pivotField.Name = "Territory";
                Console.WriteLine("Pivot field \"Region\" hidden and renamed to \"Territory\".");
            }
            else
            {
                Console.WriteLine("Pivot field \"Region\" not found.");
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
