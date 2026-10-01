// Title: How to add a Top 10 filter to a pivot table row field with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that applies a Top 10 filter to the first row field of an Aspose.Cells pivot table, showing the ten highest values in descending order. | Demonstrate using reflection to set the AutoShowTop10, Top10Count, Top10Type, and Top10SortOrder properties on a PivotField for version‑independent Aspose.Cells filtering.
// Common Searches: aspnet apply top 10 filter to pivot table row field using Aspose.Cells | c# Aspose.Cells pivot table show highest 10 items | set AutoShowTop10 property on PivotField with reflection Aspose.Cells | how to filter pivot table rows by top values in Aspose.Cells .NET | version independent top 10 filter for Aspose.Cells pivot tables
// Tags: Aspose.Cells pivot table top10 filter C# | set AutoShowTop10 via reflection Aspose.Cells | apply descending top10 filter by value PivotField | C# pivot table row field filtering Aspose.Cells | version‑agnostic pivot table filter Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, accesses the first worksheet's first pivot table, retrieves its first row field, and uses reflection to enable AutoShowTop10, set Top10Count to 10, configure Top10Type to ByValue, and set Top10SortOrder to Descending, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that a pivot table exists
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Verify that a row field exists
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("No row fields found in the pivot table.");
                return;
            }

            // Get the first row field
            PivotField rowField = pivotTable.RowFields[0];

            // Apply Top 10 filter using reflection (covers different library versions)
            try
            {
                var fieldType = rowField.GetType();

                var autoShowProp = fieldType.GetProperty("AutoShowTop10");
                var countProp = fieldType.GetProperty("Top10Count");
                var typeProp = fieldType.GetProperty("Top10Type");
                var sortOrderProp = fieldType.GetProperty("Top10SortOrder");

                if (autoShowProp != null && autoShowProp.CanWrite)
                    autoShowProp.SetValue(rowField, true);

                if (countProp != null && countProp.CanWrite)
                    countProp.SetValue(rowField, 10);

                if (typeProp != null && typeProp.CanWrite)
                {
                    var enumType = typeProp.PropertyType;
                    var enumValue = Enum.Parse(enumType, "ByValue");
                    typeProp.SetValue(rowField, enumValue);
                }

                if (sortOrderProp != null && sortOrderProp.CanWrite)
                {
                    var enumType = sortOrderProp.PropertyType;
                    var enumValue = Enum.Parse(enumType, "Descending");
                    sortOrderProp.SetValue(rowField, enumValue);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying Top 10 filter: {ex.Message}");
                // Continue without the filter if it fails
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
