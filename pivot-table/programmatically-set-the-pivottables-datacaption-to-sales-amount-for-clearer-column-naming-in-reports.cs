// Title: C# example: set the caption of a PivotTable data field to "Sales Amount" using Aspose.Cells
// Common Searches: Aspose.Cells C# change pivot table data field caption to custom text | set pivot table value field name programmatically in .NET with Aspose.Cells | rename first data field of an Excel PivotTable to Sales Amount using Aspose API | how to modify pivot table column heading before saving workbook in C# | update pivot table data field caption in an existing workbook with Aspose.Cells
// Tags: set pivot table data field name Aspose.Cells | modify pivot table column heading C# | rename pivot table value field programmatically | update Excel pivot table caption using Aspose | Aspose.Cells pivot table field caption change

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing Excel workbook, accesses the first worksheet's PivotTable, and if a data field exists, assigns the name "Sales Amount" to the first data field to change its column caption, then saves the workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet where the PivotTable resides (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Set a clear column heading for the first data field, if any
            if (pivotTable.DataFields.Count > 0)
            {
                // The Name property of a PivotField defines its caption in the PivotTable
                pivotTable.DataFields[0].Name = "Sales Amount";
            }
            else
            {
                Console.WriteLine("PivotTable has no data fields to rename.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
