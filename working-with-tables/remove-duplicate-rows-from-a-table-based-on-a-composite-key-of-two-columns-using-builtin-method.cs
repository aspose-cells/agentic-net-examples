// Title: How to remove duplicate rows from an Excel table using Aspose.Cells RemoveDuplicates with a composite key in C#
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, selects the first ListObject, and calls Worksheet.Cells.RemoveDuplicates using column indexes 0 and 1 as a composite key. | Adjust the duplicate‑removal logic to keep the last occurrence of each composite key instead of the first, while using the Aspose.Cells RemoveDuplicates method. | Create a command‑line version of the program that accepts a comma‑separated list of column indexes for the composite key and processes any input Excel file.
// Common Searches: Aspose.Cells C# remove duplicate rows based on two columns in a ListObject | Worksheet.Cells.RemoveDuplicates example with composite key in C# | How to keep the first row while deleting duplicate records in an Excel table using Aspose.Cells | C# code to delete duplicate rows in an Excel table by columns A and B with Aspose.Cells | Remove duplicate records from an Excel table using Aspose.Cells RemoveDuplicates method
// Tags: aspocells removeduplicates method | excel table composite key deduplication | listobject duplicate elimination c# | worksheet cells deduplication case-sensitive | c# aspocells duplicate row removal

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an input.xlsx workbook, accesses the first worksheet and its first ListObject, determines the data range, and calls Worksheet.Cells.RemoveDuplicates with a composite key consisting of the first two columns (A and B). It removes duplicate rows while preserving the first occurrence and saves the cleaned workbook as output.xlsx.
class RemoveDuplicateRows
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

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Define zero‑based column indexes that form the composite key (e.g., columns A and B)
            int[] keyColumns = new int[] { 0, 1 };

            // Determine the data range of the table
            Aspose.Cells.Range dataRange = table.DataRange;
            int startRow = dataRange.FirstRow;
            int startColumn = dataRange.FirstColumn;
            int totalRows = dataRange.RowCount;
            int totalColumns = dataRange.ColumnCount;

            // Remove duplicate rows based on the composite key, keeping the first occurrence
            // Note: In recent Aspose.Cells versions the parameter order is (isCaseSensitive, keyColumns)
            sheet.Cells.RemoveDuplicates(startRow, startColumn, totalRows, totalColumns, true, keyColumns);

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
