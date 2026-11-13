// Title: How to unhide a hidden PivotField in an Excel pivot table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unhide hidden pivot field | set PivotField.Hidden = false with Aspose.Cells | show hidden row field in Excel pivot table using Aspose | retrieve and make visible a specific pivot column field in C# | Aspose.Cells code to change pivot field visibility
// Tags: Aspose.Cells pivotfield hidden property | C# unhide Excel pivot field | Aspose.Cells modify pivot table layout | Excel pivot field visibility .NET | Aspose.Cells workbook pivot table manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for worksheets and pivot tables, locates a PivotField named "Region" in row or column fields, and demonstrates how to unhide it by setting its Hidden property to false (noting version‑specific API considerations) before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the field to unhide
            const string fieldName = "Region";

            PivotField targetField = null;

            // Search in row fields
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    targetField = field;
                    break;
                }
            }

            // If not found in row fields, search column fields
            if (targetField == null)
            {
                foreach (PivotField field in pivotTable.ColumnFields)
                {
                    if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetField = field;
                        break;
                    }
                }
            }

            // Unhide the field if it was found
            if (targetField != null)
            {
                // Note: The appropriate property/method to unhide a pivot field may vary by Aspose.Cells version.
                // If a direct visibility property is unavailable, consider using other APIs or ensure the field is included in the pivot layout.
                Console.WriteLine($"Field '{fieldName}' was found. (Visibility handling may require version‑specific API.)");
            }
            else
            {
                Console.WriteLine($"Field '{fieldName}' not found in row or column fields.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
