// Title: Replace an Excel QueryTable with a regular ListObject table while preserving data and formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert QueryTable to ListObject preserving formatting | replace Excel query table with worksheet table using Aspose.Cells .NET | remove QueryTable after creating ListObject in Aspose.Cells | add ListObject with headers to existing data range Aspose.Cells | keep cell styles when converting query table to regular table Aspose.Cells
// Tags: convert QueryTable to ListObject Aspose.Cells | preserve formatting when adding ListObject | remove QueryTable definition Aspose.Cells | apply built‑in table style Aspose.Cells | create worksheet table from used range C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, checks for a QueryTable on the first worksheet, determines the used cell range, adds a ListObject (regular table) over that range with headers and a built‑in style, deletes the original QueryTable definition, and saves the updated workbook.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the query table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Proceed only if there is at least one query table
            if (worksheet.QueryTables.Count > 0)
            {
                // Get the first query table
                QueryTable queryTable = worksheet.QueryTables[0];

                // Determine the used range of the worksheet as an approximation of the query table data range
                int lastRow = worksheet.Cells.MaxDataRow;
                int lastCol = worksheet.Cells.MaxDataColumn;

                // Ensure the worksheet is not empty
                if (lastRow >= 0 && lastCol >= 0)
                {
                    // Calculate range dimensions (Aspose.Cells uses zero‑based indices)
                    int firstRow = 0;
                    int firstColumn = 0;
                    int totalRows = lastRow - firstRow + 1;
                    int totalColumns = lastCol - firstColumn + 1;

                    // Add a ListObject (regular table) over the same range.
                    // The overload requires row/column indices and a hasHeaders flag.
                    int listObjectIndex = worksheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                    ListObject listObject = worksheet.ListObjects[listObjectIndex];

                    // Apply a built‑in table style (optional)
                    listObject.TableStyleType = TableStyleType.TableStyleMedium2;
                }

                // Remove the original query table definition
                worksheet.QueryTables.RemoveAt(0);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the program from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
