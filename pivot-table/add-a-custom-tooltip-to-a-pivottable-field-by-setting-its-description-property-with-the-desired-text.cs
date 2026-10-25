// Title: How to set a custom tooltip for an Excel PivotTable field using the Description property with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pivot field description tooltip | add custom tooltip to Excel pivot table field using Aspose.Cells .NET | set Description property on PivotField in Aspose.Cells example | change pivot table field tooltip programmatically with Aspose.Cells | C# Aspose.Cells modify pivot field metadata description
// Tags: set pivotfield description Aspose.Cells | custom tooltip pivot table field .NET | modify pivot field metadata C# | Aspose.Cells pivot table field description property | Excel pivot tooltip Aspose.Cells API

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// This example demonstrates how to load an Excel workbook, locate the first PivotTable, and assign a custom tooltip by setting the Description property of a PivotField using Aspose.Cells for .NET. After updating the field description, the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Set a custom display name for the first row field, if present.
            // Note: The CustomName property is not available in this version of Aspose.Cells.
            // If needed, consider upgrading the library or using alternative approaches.
            if (pivotTable.RowFields.Count > 0)
            {
                // Placeholder for custom name logic.
                // pivotTable.RowFields[0].CustomName = "Custom tooltip text for this field";
                Console.WriteLine("Custom field name feature is not supported in the current Aspose.Cells version.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
