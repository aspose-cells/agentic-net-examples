// Title: Detect Excel worksheets that contain only column headers with Aspose.Cells for .NET (MaxDataRow = -1 and MaxDataColumn > 0)
// Common Searches: how to find Excel sheets that have only header rows using Aspose.Cells in C# | Aspose.Cells MaxDataRow -1 MaxDataColumn >0 to identify header‑only worksheets | C# code to detect worksheets without data rows but with column headers | determine if an Excel worksheet contains only column headers using Aspose.Cells | list worksheets that contain only column headers in a workbook with Aspose.Cells .NET
// Tags: header-only worksheet detection Aspose.Cells | maxdatarow -1 maxdatacolumn >0 check C# | excel sheet column header identification .NET | worksheet data presence analysis Aspose.Cells | detect empty data rows with column headers in Excel

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and uses Cells.MaxDataRow and Cells.MaxDataColumn to determine whether a sheet contains only column headers (MaxDataRow == -1 and MaxDataColumn > 0). It reports header‑only worksheets and those with data rows to the console.
class Program
{
    static void Main(string[] args)
    {
        // Expect the first argument to be the path of the Excel file to analyze.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <path-to-excel-file>");
            return;
        }

        string filePath = args[0];

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at \"{filePath}\"");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine($"Workbook \"{Path.GetFileName(filePath)}\" loaded successfully.");
            bool foundHeaderOnlySheet = false;

            // Iterate through all worksheets.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i];
                Cells cells = worksheet.Cells;

                // Retrieve the maximum data row and column indices.
                int maxDataRow = cells.MaxDataRow;       // Returns -1 if no data cell exists.
                int maxDataColumn = cells.MaxDataColumn; // Returns -1 if no data cell exists.

                // According to the task, a sheet with only column headers has:
                //   MaxDataRow == -1  &&  MaxDataColumn > 0
                if (maxDataRow == -1 && maxDataColumn > 0)
                {
                    Console.WriteLine($"[HeaderOnly] Worksheet \"{worksheet.Name}\" (Index {i}) contains only column headers.");
                    foundHeaderOnlySheet = true;
                }
                else
                {
                    Console.WriteLine($"[DataPresent] Worksheet \"{worksheet.Name}\" (Index {i}) - MaxDataRow: {maxDataRow}, MaxDataColumn: {maxDataColumn}");
                }
            }

            if (!foundHeaderOnlySheet)
            {
                Console.WriteLine("No worksheets with only column headers were detected.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
