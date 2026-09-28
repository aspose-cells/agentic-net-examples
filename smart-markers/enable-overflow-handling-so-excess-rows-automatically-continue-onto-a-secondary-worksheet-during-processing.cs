// Title: Import a CSV with over 65,000 rows and automatically split overflow rows onto additional worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a CSV containing 70,000 rows into an Aspose.Cells Workbook and creates extra worksheets when the Excel row limit is reached by setting TxtLoadOptions.ExtendToNextSheet. | Show how to enable overflow handling for CSV imports in Aspose.Cells so that rows beyond 65,535 are placed on a new sheet automatically.
// Common Searches: asp.net aspose.cells load csv exceeding 65535 rows | c# TxtLoadOptions ExtendToNextSheet example | how to split large CSV into multiple worksheets with Aspose.Cells | automatic worksheet creation when CSV row count exceeds Excel limit | aspose.cells overflow handling for CSV import
// Tags: CSV import overflow handling Aspose.Cells | ExtendToNextSheet TxtLoadOptions | automatic worksheet creation for large CSV | exceed Excel row limit Aspose.Cells .NET | split CSV rows across multiple sheets

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace OverflowHandlingDemo
{
    // The example generates a 70,000‑row CSV in memory, loads it into an Aspose.Cells Workbook with TxtLoadOptions.ExtendToNextSheet enabled, causing rows beyond the 65,535 limit to be placed on a second worksheet, then reports the worksheet count and saves the file as OverflowDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Generate CSV data with more rows than the Excel 2003 limit (65535 rows)
            const int totalRows = 70000; // exceeds the limit to trigger overflow
            StringBuilder csvBuilder = new StringBuilder();

            // Header row
            csvBuilder.AppendLine("Index,Value");

            // Data rows
            for (int i = 1; i <= totalRows; i++)
            {
                csvBuilder.AppendLine($"{i},Data_{i}");
            }

            // Convert CSV string to a memory stream
            using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvBuilder.ToString())))
            {
                // Enable overflow handling so excess rows go to the next worksheet
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    ExtendToNextSheet = true
                };

                // Load the CSV data into a workbook using the specified options
                Workbook workbook = new Workbook(csvStream, loadOptions);

                // Output information about the resulting workbook
                Console.WriteLine($"Number of worksheets created: {workbook.Worksheets.Count}");
                Console.WriteLine($"Rows in first worksheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
                if (workbook.Worksheets.Count > 1)
                {
                    Console.WriteLine($"Rows in second worksheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");
                }

                // Save the workbook to a file
                workbook.Save("OverflowDemo.xlsx");
                Console.WriteLine("Workbook saved as OverflowDemo.xlsx");
            }
        }
    }
}
