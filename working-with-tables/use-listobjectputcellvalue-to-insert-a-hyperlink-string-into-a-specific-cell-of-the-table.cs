// Title: Insert a hyperlink into a specific cell of an Aspose.Cells ListObject table using C#
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject table, and uses ListObject.PutCellValue to set a cell to a URL string. | Write a C# snippet that places the hyperlink "https://www.example.com" into the third row, second column of an Aspose.Cells ListObject and saves the file. | Provide C# example code for adding a ListObject to a worksheet, inserting a hyperlink into a table cell with PutCellValue, and handling output folder creation.
// Common Searches: Aspose.Cells C# put hyperlink into ListObject table cell | How to use ListObject.PutCellValue to add a URL in Aspose.Cells | C# Aspose.Cells insert URL into specific table cell example | Saving workbook after adding hyperlink to Excel table with Aspose.Cells
// Tags: Aspose.Cells ListObject PutCellValue hyperlink | C# insert URL into Excel table cell | Aspose.Cells create ListObject table | save workbook Aspose.Cells C# | set cell value to string Aspose.Cells | handle output directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a ListObject spanning A1:C5, and uses ListObject.PutCellValue to insert the URL https://www.example.com into the cell at the third row, second column of the table, then saves the workbook to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range for the table (ListObject)
                // Table starts at cell A1 and spans 5 rows × 3 columns
                int firstRow = 0;          // zero‑based index for row 1 (A1)
                int firstColumn = 0;       // zero‑based index for column A
                int totalRows = 5;
                int totalColumns = 3;

                // Add the ListObject (table) to the worksheet
                int tableIndex = worksheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows,
                    firstColumn + totalColumns,
                    true);

                ListObject table = worksheet.ListObjects[tableIndex];

                // Insert a hyperlink string into a specific cell of the table
                // 3rd row (index 2) and 2nd column (index 1) within the table
                string hyperlink = "https://www.example.com";
                table.PutCellValue(2, 1, hyperlink);

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory exists (handle case when outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
