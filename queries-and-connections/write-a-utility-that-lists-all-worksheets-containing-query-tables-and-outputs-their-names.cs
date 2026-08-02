using System;
using Aspose.Cells;

namespace QueryTableUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the source workbook path (first argument or default)
            string sourcePath = args.Length > 0 ? args[0] : "input.xlsx";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet contains any query tables
                if (sheet.QueryTables != null && sheet.QueryTables.Count > 0)
                {
                    // Output the worksheet name
                    Console.WriteLine($"Worksheet with QueryTables: {sheet.Name}");
                }
            }

            // Save the workbook (save rule) – no changes made, just demonstrating the rule usage
            workbook.Save("output.xlsx");
        }
    }
}