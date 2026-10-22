// Title: Unhide a hidden column field in an Excel PivotTable using Aspose.Cells for C#
// Common Searches: C# how to unhide a pivot table column field using Aspose.Cells | Aspose.Cells example for setting PivotField IsHidden to false | programmatically show hidden column field in Excel pivot table with C# | toggle visibility of a pivot column field by name in Aspose.Cells | unhide pivot column field by index using Aspose.Cells C# code
// Tags: Aspose.Cells PivotField IsHidden property | C# set pivot column field visibility | Excel pivot table column field visibility control | dynamic field visibility handling Aspose.Cells | programmatic pivot table field visibility C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a PivotTable on the first worksheet, locates a column field by name (or index), uses a dynamic cast to set its IsHidden property to false (with version‑specific fallback), and saves the workbook with the previously hidden column field now visible.
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

            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a PivotTable exists on the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found on the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];
            PivotFieldCollection columnFields = pivotTable.ColumnFields;

            // Example 1: Show a hidden column field by its name
            string hiddenFieldName = "FieldName"; // replace with the actual field name
            foreach (PivotField field in columnFields)
            {
                if (field.Name == hiddenFieldName)
                {
                    try
                    {
                        // Use dynamic to access IsHidden property (available in newer versions)
                        dynamic dynField = field;
                        dynField.IsHidden = false;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // Property not available in this version; alternative handling can be added here
                        Console.WriteLine($"Unable to set visibility for field \"{hiddenFieldName}\" – IsHidden property not supported.");
                    }
                    break;
                }
            }

            // Example 2: If you know the index of the hidden column field, you can set it directly
            // int hiddenFieldIndex = 0; // replace with the correct zero‑based index
            // try
            // {
            //     dynamic dynField = columnFields[hiddenFieldIndex];
            //     dynField.IsHidden = false;
            // }
            // catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
