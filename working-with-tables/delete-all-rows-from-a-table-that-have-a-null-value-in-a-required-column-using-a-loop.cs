// Title: Delete rows with null values in a required column from an Excel ListObject using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, locates a ListObject by name, and removes every row where a specified column contains a null or empty string, then saves the workbook. | Write a reverse‑order loop in C# using Aspose.Cells to safely eliminate worksheet rows that have missing data in a required column of a table. | Adapt the example to eliminate rows based on multiple required columns while keeping the ListObject intact with Aspose.Cells.
// Common Searches: aspnet delete rows from Excel table where column value is null using Aspose.Cells | c# remove rows with empty cells in a ListObject with Aspose.Cells | how to iterate Excel table rows in reverse and delete rows in Aspose.Cells | Aspose.Cells delete rows based on required column condition in .NET
// Tags: Aspose.Cells null column row purge | C# ListObject reverse iteration deletion | Excel table required column validation Aspose.Cells | Aspose.Cells workbook save after row cleanup | Aspose.Cells multi‑column row removal

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// // Loads "input.xlsx", finds the ListObject "Table1", identifies the "RequiredColumn", iterates rows in reverse, deletes any worksheet row where that column is null or whitespace, and saves the result to "output.xlsx".
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the table (ListObject) by its name; replace "Table1" with your actual table name
            ListObject table = sheet.ListObjects["Table1"];
            if (table == null)
            {
                Console.WriteLine("Error: Table \"Table1\" not found in the worksheet.");
                return;
            }

            // Identify the required column by its header name; replace "RequiredColumn" with the actual column name
            int requiredColIndex = -1;
            for (int c = 0; c < table.ListColumns.Count; c++)
            {
                if (table.ListColumns[c].Name.Equals("RequiredColumn", StringComparison.OrdinalIgnoreCase))
                {
                    requiredColIndex = c;
                    break;
                }
            }

            if (requiredColIndex == -1)
            {
                Console.WriteLine("Error: Column \"RequiredColumn\" not found in the table.");
                return;
            }

            // Loop through the data rows of the table in reverse order to safely delete rows
            for (int i = table.DataRange.RowCount - 1; i >= 0; i--)
            {
                // Calculate the absolute cell position for the required column in the current row
                int rowIndex = table.DataRange.FirstRow + i;
                int colIndex = table.DataRange.FirstColumn + requiredColIndex;

                // Retrieve the cell
                Cell cell = sheet.Cells[rowIndex, colIndex];

                // Check if the cell value is null or empty
                if (cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    // Delete the entire row from the worksheet (the table will adjust automatically)
                    sheet.Cells.DeleteRow(rowIndex);
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
