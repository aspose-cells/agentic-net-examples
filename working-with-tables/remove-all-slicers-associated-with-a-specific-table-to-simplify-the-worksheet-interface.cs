// Title: Delete all slicers associated with a specific table in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans each worksheet, identifies slicers whose TableName equals a specified value, purges them, and writes the updated file. | Demonstrate how to traverse the Slicers collection in reverse order to safely eliminate matching slicers in Aspose.Cells. | Include file existence verification and per‑slicer exception handling while deleting slicers for a target table in a .NET example.
// Common Searches: aspnet remove slicers from Excel workbook by table name using Aspose.Cells | c# code to delete slicers linked to a specific table in .xlsx | how to iterate slicer collection backwards in Aspose.Cells | programmatically clear slicers for a given table with Aspose.Cells .NET | sample to delete Excel slicers for Table1 using Aspose.Cells
// Tags: Aspose.Cells remove slicers by table name | C# delete slicer collection .xlsx | iterate worksheets remove slicers Aspose.Cells | backward loop slicer removal C# | Excel slicer cleanup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an input.xlsx workbook, checks that the file exists, then iterates through each worksheet. It walks the Slicers collection in reverse order, compares each slicer’s TableName to the target name (e.g., "Table1"), and removes matching slicers. After processing all sheets, the workbook is saved as output.xlsx. The code includes robust error handling for missing files and individual slicer processing failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetTableName = "Table1";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards to safely remove slicers while iterating
                for (int i = sheet.Slicers.Count - 1; i >= 0; i--)
                {
                    // Use dynamic to avoid compile‑time dependency on the Slicer type
                    dynamic slicer = sheet.Slicers[i];
                    try
                    {
                        string slicerTableName = slicer.TableName as string;
                        if (string.Equals(slicerTableName, targetTableName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheet.Slicers.RemoveAt(i);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue if a particular slicer cannot be processed
                        Console.WriteLine($"Error processing slicer at index {i} on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
