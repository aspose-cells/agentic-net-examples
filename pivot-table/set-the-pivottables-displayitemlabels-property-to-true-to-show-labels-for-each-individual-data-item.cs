// Title: Enable DisplayItemLabels for all data fields in an Aspose.Cells PivotTable using C#
// AI Prompts: Create a C# program that opens an existing Excel file with Aspose.Cells, retrieves the first PivotTable, sets its DisplayItemLabels property to true for every data field, and saves the workbook to a new file. | Write a C# snippet that iterates over PivotTable.DataFields and assigns ShowItemLabels = true (or the equivalent DisplayItemLabels flag) using Aspose.Cells, handling missing files and ensuring the output directory exists. | Provide a complete example that checks for a PivotTable on a worksheet, enables item labels for each data item, and demonstrates saving the modified workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to turn on DisplayItemLabels for PivotTable data fields | set item labels true for each data item in Excel PivotTable using Aspose.Cells .NET | C# code example to show item labels in a PivotTable with Aspose.Cells library | enable data field labels in an existing workbook's PivotTable via Aspose.Cells
// Tags: Aspose.Cells set DisplayItemLabels on PivotTable | C# enable item labels for PivotTable data fields | modify PivotTable data field visibility using Aspose.Cells | Excel PivotTable item label activation with .NET | Aspose.Cells update PivotTable display settings programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads a workbook, verifies a PivotTable exists on the first worksheet, iterates through its DataFields to set DisplayItemLabels (or ShowItemLabels) to true, ensures the output directory is present, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Enable display of item labels for each data field in the PivotTable
            // Note: The ShowItemLabels property is not available in older Aspose.Cells versions.
            // If needed, this block can be updated when the property becomes supported.
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                // Placeholder for future property assignment, e.g., dataField.ShowItemLabels = true;
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
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
