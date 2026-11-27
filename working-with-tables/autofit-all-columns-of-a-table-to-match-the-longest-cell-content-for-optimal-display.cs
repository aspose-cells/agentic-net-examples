// Title: C# example: Auto‑fit columns of an Excel ListObject table using Aspose.Cells
// Common Searches: Aspose.Cells C# auto fit columns for a ListObject table | How to adjust column width to the longest cell value in an Excel table with Aspose.Cells | C# code to auto size columns of a specific table range in a workbook using Aspose.Cells | AutoFitColumn for each column in an Excel table using Aspose.Cells .NET | Resize Excel table columns to content length programmatically with Aspose.Cells
// Tags: Aspose.Cells AutoFitColumn ListObject | C# auto size Excel table columns | adjust column width based on content Aspose.Cells | Excel table column auto‑fit .NET | worksheet column auto‑fit for table range

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a ListObject (Excel table) on the first worksheet, determines the table's full cell range, iterates through each column in that range calling AutoFitColumn to size columns according to the longest cell content, and saves the updated workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Process the first table (ListObject) if present
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first table
                ListObject table = sheet.ListObjects[0];

                // Determine the full range of the table (including header and data)
                CellArea tableArea = new CellArea
                {
                    StartRow = table.StartRow,
                    StartColumn = table.StartColumn,
                    EndRow = table.EndRow,
                    EndColumn = table.EndColumn
                };

                // Auto‑fit each column within the table range
                for (int col = tableArea.StartColumn; col <= tableArea.EndColumn; col++)
                {
                    sheet.AutoFitColumn(col);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
