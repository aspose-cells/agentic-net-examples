// Title: Refresh an existing AutoFilter on a named Excel table (ListObject) using Aspose.Cells for .NET when CriteriaRange is not supported
// AI Prompts: Load an .xlsx file, locate the ListObject called 'MyTable' on the first worksheet, invoke the method that updates its current filter, and save the workbook using Aspose.Cells in C#. | With Aspose.Cells for .NET, retrieve a table by its name, trigger a refresh of any applied filters, and write the modified workbook to a new file. | Demonstrate a workaround for the missing CriteriaRange property by simply refreshing the existing filter on a named Excel table with C# code.
// Common Searches: Aspose.Cells how to refresh autofilter on a ListObject table in C# | C# code to update existing filter on Excel table using Aspose.Cells | Example of using Aspose.Cells to refresh table filters | Workaround for missing CriteriaRange in Aspose.Cells AutoFilter | Refresh Excel table filter after loading workbook with Aspose.Cells .NET
// Tags: AutoFilter.Refresh Aspose.Cells C# | ListObject filter refresh Aspose.Cells | named table access Aspose.Cells .NET | CriteriaRange not supported Aspose.Cells | load workbook save after filter refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, retrieves the ListObject named 'MyTable', calls the AutoFilter.Refresh method to update any existing filter (since CriteriaRange is unavailable), and saves the result as 'output.xlsx', including basic file existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) by name
                ListObject table = sheet.ListObjects["MyTable"];
                if (table == null)
                {
                    Console.WriteLine("Table 'MyTable' not found.");
                    return;
                }

                // NOTE: Aspose.Cells AutoFilter does not expose a CriteriaRange property.
                // If specific filter criteria are required, use the appropriate AutoFilter methods
                // such as ApplyCustomFilter. Here we simply refresh the existing filter.
                table.AutoFilter.Refresh();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
