using System;
using System.IO;
using Aspose.Cells;

namespace ConditionalFreezeDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add sample worksheets
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Report_January";

                // Worksheets.Add(string) returns the created Worksheet
                Worksheet sheet2 = workbook.Worksheets.Add("Data_Q1");
                Worksheet sheet3 = workbook.Worksheets.Add("Report_February");

                // Iterate through all worksheets
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Apply freeze only when the worksheet name starts with "Report"
                    if (ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                    {
                        // Freeze panes at cell C3 (row index 2, column index 2) with 3 frozen rows and 3 frozen columns
                        ws.FreezePanes("C3", 3, 3);
                    }
                }

                // Define output file path
                string outputPath = "ConditionalFreezeDemo.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}