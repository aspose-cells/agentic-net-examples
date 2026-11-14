// Title: Iterate through all worksheets in an Excel workbook and convert each table to a range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate worksheets and convert tables to ranges | How to delete Excel tables while keeping formatting with Aspose.Cells | Convert ListObject to range in a workbook using Aspose.Cells .NET | Batch convert all tables to ranges in an Excel file with Aspose.Cells | Aspose.Cells ConvertToRange example for multiple worksheets
// Tags: Aspose.Cells ConvertToRange ListObject | C# iterate workbook tables Aspose.Cells | remove Excel tables preserve formatting Aspose.Cells | batch table-to-range conversion Aspose.Cells | reverse loop ListObjectCollection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject and ListObjectCollection

// The sample loads an existing XLSX file, checks its existence, then loops through every worksheet. For each sheet it retrieves the ListObjectCollection, iterates the tables in reverse order, and calls ConvertToRange() to replace each table with a regular range while preserving formatting. After processing all sheets, the workbook is saved to a new file.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of tables (ListObjects) on the current worksheet
                ListObjectCollection tables = sheet.ListObjects;

                // Iterate backwards to safely handle removal of tables during conversion
                for (int i = tables.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        ListObject table = tables[i];
                        // Convert the table to a range using default options (preserves formatting and deletes the table)
                        table.ConvertToRange();
                    }
                    catch (Exception exTable)
                    {
                        Console.WriteLine($"Failed to convert table on sheet \"{sheet.Name}\": {exTable.Message}");
                    }
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
