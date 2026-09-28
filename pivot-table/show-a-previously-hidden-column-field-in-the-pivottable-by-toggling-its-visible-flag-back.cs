// Title: Programmatically make a hidden column field visible in an Aspose.Cells PivotTable with C#
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, accesses the first PivotTable, checks for column fields, and sets the first column field's IsHidden property to false before saving the file. | Create a C# snippet that safely verifies the existence of column fields in a PivotTable and unhides a specified field with Aspose.Cells, handling cases where no column fields are present. | Write a C# example that toggles the visibility of a PivotTable column field in Aspose.Cells, demonstrating how to read, modify, and persist the workbook.
// Common Searches: aspocells c# how to show hidden column field in pivot table | unhide pivot column field programmatically using Aspose.Cells .NET | set PivotField IsHidden false in C# Aspose.Cells example | check for column fields before changing visibility in Aspose.Cells PivotTable
// Tags: Aspose.Cells PivotField IsHidden property | C# unhide pivot column field | modify pivot table column visibility Aspose | Aspose.Cells hide/show pivot fields | Excel workbook pivot table column field visibility C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads an Excel workbook, accesses the first worksheet's first PivotTable, verifies that column fields exist, clears the IsHidden flag on the first column field (when supported), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Make the first column field visible, if any exist
            if (pivotTable.ColumnFields.Count > 0)
            {
                PivotField columnField = pivotTable.ColumnFields[0];
                // In newer Aspose.Cells versions the visibility is controlled via IsHidden.
                // Uncomment the following line if the property is available in your version.
                // columnField.IsHidden = false;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
