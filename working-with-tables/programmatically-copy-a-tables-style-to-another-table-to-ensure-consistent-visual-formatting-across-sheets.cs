// Title: Copy an Excel table’s style from one worksheet to another with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy table formatting from Sheet1 to Sheet2 | how to transfer ListObject TableStyleName between worksheets using Aspose.Cells | programmatically apply the same Excel table style to another table in .NET | Aspose.Cells set destination table style to match source table | duplicate Excel table style across sheets with Aspose.Cells API
// Tags: Aspose.Cells copy table style C# | ListObject TableStyleName assignment Aspose | duplicate Excel table formatting .NET | apply consistent table style across worksheets | transfer table style between worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads a workbook, reads the TableStyleName from the first table on Sheet1, assigns that style to the first table on Sheet2, and saves the updated workbook, ensuring both tables share identical visual formatting.
class CopyTableStyle
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get source worksheet and its first table
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Sheet1' not found.");
                return;
            }

            ListObjectCollection sourceTables = sourceSheet.ListObjects;
            if (sourceTables == null || sourceTables.Count == 0)
            {
                Console.WriteLine("No tables found in source worksheet.");
                return;
            }
            ListObject sourceTable = sourceTables[0];

            // Retrieve the style name applied to the source table
            string sourceStyleName = sourceTable.TableStyleName;

            // Get destination worksheet and its first table
            Worksheet destSheet = workbook.Worksheets["Sheet2"];
            if (destSheet == null)
            {
                Console.WriteLine("Destination worksheet 'Sheet2' not found.");
                return;
            }

            ListObjectCollection destTables = destSheet.ListObjects;
            if (destTables == null || destTables.Count == 0)
            {
                Console.WriteLine("No tables found in destination worksheet.");
                return;
            }
            ListObject destTable = destTables[0];

            // Apply the source table's style to the destination table
            destTable.TableStyleName = sourceStyleName;

            // Save the workbook with the updated style
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
