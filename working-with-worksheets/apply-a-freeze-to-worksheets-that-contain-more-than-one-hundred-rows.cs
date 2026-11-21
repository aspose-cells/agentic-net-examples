// Title: How to freeze the first 100 rows in Excel worksheets that contain more than 100 rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze top 100 rows only when worksheet has more than 100 rows | C# apply FreezePanes to each sheet of an Excel workbook if row count exceeds 100 using Aspose.Cells | detect max data row and conditionally set freeze panes in Aspose.Cells .NET | skip freezing panes on sheets with fewer than 100 rows Aspose.Cells C# example | iterate all worksheets and freeze first hundred rows in large sheets with Aspose.Cells
// Tags: freeze panes based on row count Aspose.Cells | Aspose.Cells FreezePanes for Excel worksheets | detect max data row in worksheet C# | iterate workbook worksheets Aspose.Cells | conditional freeze panes .NET Excel

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, loops through each worksheet, checks the last used row with Cells.MaxDataRow, and if a sheet has more than 100 rows it freezes the first 100 rows (no columns) using the FreezePanes(100, 0, 100, 0) overload, then saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used row (0‑based). Returns -1 if the sheet is empty.
                int lastUsedRow = sheet.Cells.MaxDataRow;

                // If the sheet contains more than 100 rows (i.e., row index >= 100)
                if (lastUsedRow + 1 > 100)
                {
                    // Freeze the first 100 rows. Columns are not frozen.
                    // Use the 4‑parameter overload required by recent Aspose.Cells versions.
                    sheet.FreezePanes(100, 0, 100, 0);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
