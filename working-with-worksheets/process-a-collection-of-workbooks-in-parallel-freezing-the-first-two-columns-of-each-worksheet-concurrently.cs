// Title: Parallel freeze of the first two columns in all worksheets using Aspose.Cells for .NET (C#)
// Description: Loads a collection of Excel files, processes each workbook concurrently with Parallel.ForEach, iterates every worksheet, applies FreezePanes at cell C1 to lock the first two columns, and saves the changes back to the original files.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze first two columns | parallel processing | batch Excel | multiple workbooks | concurrent worksheet freeze | Parallel.ForEach | Excel automation
// Common Searches: freeze first two columns in all sheets Aspose.Cells C# | parallel processing of multiple Excel workbooks Aspose.Cells | batch freeze panes Aspose.Cells .NET | how to use FreezePanes with Parallel.ForEach | Aspose.Cells example for freezing columns across workbooks
// Developer Intent: Apply column freezing to the first two columns of every worksheet in a set of workbooks while executing the operation in parallel.
// Use Cases: Prepare a fleet of financial report workbooks for distribution by locking navigation columns in one batch step. | Accelerate preprocessing of template‑based Excel outputs in an ETL pipeline by freezing panes concurrently. | Integrate a parallel column‑freeze routine into a CI/CD workflow that validates and formats generated spreadsheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze the first three rows and the first column of each worksheet in a list of workbooks processed in parallel. | Show how to add robust exception handling for loading and saving workbooks inside a Parallel.ForEach loop with Aspose.Cells. | Provide an example that freezes both rows and columns (e.g., first two rows and first two columns) for every worksheet while maintaining parallel execution.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads a collection of Excel files, processes each workbook concurrently with Parallel.ForEach, iterates every worksheet, applies FreezePanes at cell C1 to lock the first two columns, and saves the changes back to the original files.
class ParallelFreezePanes
{
    static void Main()
    {
        // List of workbook file paths to process
        List<string> workbookFiles = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
            // Add more file paths as needed
        };

        // Process each workbook in parallel
        Parallel.ForEach(workbookFiles, filePath =>
        {
            // Load the workbook (uses the provided Workbook(string) constructor)
            using (Workbook workbook = new Workbook(filePath))
            {
                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Freeze the first two columns (C1 is the split point, 0 rows frozen, 2 columns frozen)
                    sheet.FreezePanes("C1", 0, 2);
                }

                // Save the modified workbook (uses the provided Save(string) method)
                workbook.Save(filePath);
            }
        });

        Console.WriteLine("All workbooks have been processed and first two columns frozen.");
    }
}
