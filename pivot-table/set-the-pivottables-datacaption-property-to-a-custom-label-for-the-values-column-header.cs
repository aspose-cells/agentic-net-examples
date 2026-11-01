// Title: Set a custom caption for the values column of a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to change the caption of the first data field in an existing PivotTable to a custom label and save the workbook. | Programmatically rename the values column header of a PivotTable by assigning a new Name to its DataFields[0] with Aspose.Cells for .NET. | Load an Excel file, modify the PivotTable's data field caption to "Custom Values", and write the updated file using Aspose.Cells C# API.
// Common Searches: aspocells change pivot table values column header c# | rename data field caption in Excel pivot table using Aspose.Cells .NET | set custom DataCaption for pivot table values column programmatically | how to modify pivot table data field name with Aspose.Cells C#
// Tags: Aspose.Cells pivot table values column caption | C# rename pivot table data field | set DataCaption property Aspose.Cells | custom values header Excel pivot table .NET | modify PivotTable DataFields name Aspose

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads an Excel workbook, accesses the first worksheet's first PivotTable, changes the caption of its first data field to "Custom Values", and saves the modified workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Set a custom label for the first values column header, if a data field exists
            if (pivotTable.DataFields.Count > 0)
            {
                // Change the caption of the first data field
                pivotTable.DataFields[0].Name = "Custom Values";
            }
            else
            {
                Console.WriteLine("Pivot table has no data fields to rename.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
