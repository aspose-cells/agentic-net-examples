// Title: Add a default header row with generic column names to Excel tables missing headers using Aspose.Cells for .NET (C#)
// Common Searches: how to add a header row to an Excel ListObject that has no headers using Aspose.Cells C# | detect Excel tables without a header row and assign generic column names with Aspose.Cells | C# code to enable ShowHeaderRow for tables in a workbook using Aspose.Cells | automatically generate Column1 Column2 names for tables lacking headers in .xlsx files | Aspose.Cells example for iterating worksheets and adding missing table headers programmatically
// Tags: add missing table header Aspose.Cells C# | identify tables lacking header row Aspose.Cells | set ShowHeaderRow property programmatically | generate generic column names for Excel tables | populate default headers in .xlsx using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an input XLSX workbook, loops through each worksheet and its ListObject tables, checks the ShowHeaderRow flag, and for tables where it is false enables the header row and assigns generic names (Column1, Column2, …) to each column, writing those names into the header cells before saving the modified workbook.
class TableHeaderAdder
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{Path.GetFullPath(inputPath)}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all tables (ListObjects) in the worksheet
                foreach (ListObject table in sheet.ListObjects)
                {
                    // If the table does not have a header row, add one
                    if (!table.ShowHeaderRow)
                    {
                        table.ShowHeaderRow = true;

                        int headerRow = table.StartRow;
                        int startColumn = table.StartColumn;

                        // Generate generic column names and assign them
                        for (int i = 0; i < table.ListColumns.Count; i++)
                        {
                            string columnName = $"Column{i + 1}";
                            table.ListColumns[i].Name = columnName;
                            sheet.Cells[headerRow, startColumn + i].PutValue(columnName);
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
