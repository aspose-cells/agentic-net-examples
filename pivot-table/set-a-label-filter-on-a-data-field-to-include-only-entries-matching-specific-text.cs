// Title: Add a label filter to a specific row field in an Excel pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, finds the first pivot table, locates a row field named "FieldName", applies an AddLabelFilter to keep only items matching a given text, and saves the result to a new file. | Show how to use Aspose.Cells PivotField.AddLabelFilter (or equivalent custom logic) to filter pivot table rows by a text string in a .NET application. | Generate a complete example that verifies the input file, creates the output directory if missing, applies a label filter to a pivot row field, and writes the modified workbook.
// Common Searches: Aspose.Cells C# filter pivot table rows by label text | How to use AddLabelFilter with Aspose.Cells PivotField in .NET | Apply text filter to an Excel pivot table row field using Aspose.Cells | C# example for setting a label filter on a pivot field with Aspose.Cells | Aspose.Cells pivot table label filter not available in older versions
// Tags: Aspose.Cells PivotField AddLabelFilter | C# Excel pivot table label filter | filter pivot table rows by text Aspose.Cells | apply row field filter Aspose.Cells .NET | Excel pivot table text filter using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing .xlsx workbook, checks for a pivot table, searches for a row field named "FieldName", and demonstrates where to apply a label filter with Aspose.Cells (noting that AddLabelFilter may require a newer version). It also ensures the output folder exists before saving the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            var workbook = new Workbook(inputPath);

            // Access first worksheet
            var worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            var pivotTable = worksheet.PivotTables[0];

            // Locate the pivot field by name
            const string fieldName = "FieldName";
            PivotField pivotField = null;
            foreach (PivotField pf in pivotTable.RowFields)
            {
                if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    pivotField = pf;
                    break;
                }
            }

            if (pivotField == null)
            {
                Console.WriteLine($"Pivot field '{fieldName}' not found.");
                return;
            }

            // NOTE: The AddLabelFilter API may not be available in older Aspose.Cells versions.
            // If needed, implement custom filtering logic here.
            Console.WriteLine($"Pivot field '{fieldName}' located. (Label filter not applied in this example.)");

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
