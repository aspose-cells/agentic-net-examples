// Title: How to apply a custom currency number format to a column in an Excel table using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, finds the first ListObject, locates the column named "Amount", creates a style with the custom format "$#,##0.00", applies it to each cell in that column, and saves the workbook. | Show a step‑by‑step example of using Aspose.Cells Style.Custom to set a currency format for a specific table column identified by its header text. | Generate a reusable method that takes a workbook path, table name, and column header, then formats that column as currency with Aspose.Cells and writes the updated file.
// Common Searches: Aspose.Cells C# format table column as currency custom number format | Set custom number format for ListObject column in .NET Excel workbook | Apply currency style to Excel table column using Aspose.Cells API
// Tags: Aspose.Cells custom currency format | C# ListObject column styling | Excel table column number format Aspose | Apply Style.Custom to table data range | Aspose.Cells SetStyle column

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, accesses the first ListObject table, finds the "Amount" column by header, creates a style with the custom currency format "$#,##0.00", applies that style to each cell in the column's data range, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the table
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Retrieve the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Identify the column to format by its header name, e.g., "Amount"
                int tableColumnIndex = -1;
                for (int i = 0; i < table.ListColumns.Count; i++)
                {
                    if (table.ListColumns[i].Name == "Amount")
                    {
                        tableColumnIndex = i;
                        break;
                    }
                }

                if (tableColumnIndex != -1)
                {
                    // Convert the table column index to the worksheet column index
                    int worksheetColumnIndex = table.StartColumn + tableColumnIndex;

                    // Create a style with a custom currency number format
                    Style currencyStyle = workbook.CreateStyle();
                    currencyStyle.Custom = "$#,##0.00";

                    // Get the data range of the table (excludes header row)
                    AsposeRange dataRange = table.DataRange;
                    int firstDataRow = dataRange.FirstRow;
                    int lastDataRow = dataRange.FirstRow + dataRange.RowCount - 1;

                    // Apply the custom style to each cell in the target column
                    for (int row = firstDataRow; row <= lastDataRow; row++)
                    {
                        Cell cell = sheet.Cells[row, worksheetColumnIndex];
                        cell.SetStyle(currencyStyle);
                    }
                }
                else
                {
                    Console.WriteLine("Column \"Amount\" not found in the table.");
                }
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
