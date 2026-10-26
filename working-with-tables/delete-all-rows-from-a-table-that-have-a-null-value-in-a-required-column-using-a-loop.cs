// Title: Delete rows with null values in a required column from an Excel ListObject using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete rows where a specific column is empty in a ListObject | Remove rows with blank required column from an Excel table using Aspose.Cells .NET | Iterate backwards to safely delete Excel table rows with null cells in Aspose.Cells | C# code to filter out rows with null values in a required column in an Excel worksheet | Aspose.Cells delete rows based on column content without affecting table range
// Tags: delete rows based on column value Aspose.Cells | Aspose.Cells ListObject row removal | C# Excel table null column filter | backward iteration row deletion Aspose.Cells | required column null check Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads a workbook, locates a ListObject named "Table1", finds the index of a required column, iterates the data rows in reverse, deletes any row whose cell in that column is null or whitespace, and saves the cleaned workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the table (ListObject) by its name – replace "Table1" with your table's name
            ListObject table = worksheet.ListObjects["Table1"];
            if (table == null)
            {
                Console.WriteLine("Table 'Table1' not found in the worksheet.");
                return;
            }

            // Find the position of the required column inside the table
            const string requiredColumnName = "RequiredColumn";
            int requiredColumnPosition = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (string.Equals(table.ListColumns[i].Name, requiredColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    requiredColumnPosition = i;
                    break;
                }
            }

            if (requiredColumnPosition == -1)
            {
                Console.WriteLine($"Column '{requiredColumnName}' not found in the table.");
                return;
            }

            // Absolute column index in the worksheet
            int requiredColumnIndex = table.StartColumn + requiredColumnPosition;

            // Determine the data rows range (excluding header)
            int dataStartRow = table.StartRow + 1;
            int dataEndRow = table.EndRow;

            // Iterate backwards to safely remove rows
            for (int row = dataEndRow; row >= dataStartRow; row--)
            {
                Cell cell = worksheet.Cells[row, requiredColumnIndex];
                bool isEmpty = cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue);

                if (isEmpty)
                {
                    // Remove the entire row; the table range updates automatically
                    worksheet.Cells.DeleteRow(row);
                }
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
