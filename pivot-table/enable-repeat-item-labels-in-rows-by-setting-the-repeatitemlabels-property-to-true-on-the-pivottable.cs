// Title: Enable repeat item labels for row fields in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to repeat row item labels in an Excel pivot table | set IsRepeatItemLabels to true for pivot table row fields with Aspose.Cells | Aspose.Cells .NET repeat item labels for each row in a pivot table | programmatically enable repeat labels on pivot rows using Aspose.Cells | Excel pivot table repeat row labels C# Aspose.Cells example
// Tags: Aspose.Cells repeat item labels pivot row fields | C# IsRepeatItemLabels property usage | Excel pivot table row label repetition with Aspose | modify pivot table row fields Aspose.Cells | save workbook after updating pivot table labels

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a pivot table on the first worksheet, iterates through all row fields of the first pivot table, sets each field's IsRepeatItemLabels property to true to repeat item labels, and saves the modified workbook to a new file.
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

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable repeat item labels for all row fields in the pivot table
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                // Correct property name for Aspose.Cells: IsRepeatItemLabels
                rowField.IsRepeatItemLabels = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
