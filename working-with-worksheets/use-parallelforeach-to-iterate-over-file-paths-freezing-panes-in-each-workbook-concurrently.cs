using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsParallelFreeze
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example list of Excel file paths to process.
            // Replace with actual paths as needed.
            List<string> excelFiles = new List<string>
            {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xlsx",
                @"C:\Data\Report3.xlsx"
            };

            // Process each workbook in parallel.
            Parallel.ForEach(excelFiles, filePath =>
            {
                try
                {
                    // Load the workbook from the file.
                    Workbook workbook = new Workbook(filePath);

                    // Access the first worksheet (index 0).
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Freeze panes at cell "C3" with 3 frozen rows and 3 frozen columns.
                    worksheet.FreezePanes("C3", 3, 3);

                    // Save the workbook back to the same file (overwrites original).
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    // Log any errors for the specific file.
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            });

            Console.WriteLine("All workbooks have been processed.");
        }
    }
}