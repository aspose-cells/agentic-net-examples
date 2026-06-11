using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

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
            // Load the workbook (uses the Workbook(string) constructor rule)
            Workbook wb = new Workbook(filePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in wb.Worksheets)
            {
                // Freeze the first two columns.
                // Freeze at cell "C1" (column index 2) with 0 frozen rows and 2 frozen columns.
                // This uses the FreezePanes(string, int, int) method rule.
                ws.FreezePanes("C1", 0, 2);
            }

            // Save the modified workbook (uses the Workbook.Save(string) rule)
            // Overwrite the original file or specify a new path.
            wb.Save(filePath);
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}