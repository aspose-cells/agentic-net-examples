// Title: Convert every Excel table to a regular range in a workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, iterates all worksheets, and converts each ListObject table to a range with ConvertToRange while handling exceptions. | Show how to loop through a ListObjectCollection in reverse order to safely call ConvertToRange and avoid collection modification errors in Aspose.Cells. | Provide a complete example that saves the workbook after all tables have been transformed into regular ranges to a new .xlsx file.
// Common Searches: Aspose.Cells C# iterate all worksheets and convert tables to ranges | How to remove Excel tables and keep data using Aspose.Cells ConvertToRange | Convert ListObject to range in bulk with Aspose.Cells .NET | C# code to change Excel tables to normal cells using Aspose.Cells | Aspose.Cells example for converting multiple tables to ranges in one workbook
// Tags: Aspose.Cells ConvertToRange bulk tables | C# Aspose.Cells ListObject to range conversion | Excel table to range Aspose.Cells .NET | Iterate worksheets Aspose.Cells table removal | Reverse iteration ListObjectCollection Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace TableToRangeConversion
{
    // The sample verifies the input file, loads it with Aspose.Cells, iterates each worksheet and its ListObject tables in reverse order, converts every table to a regular range using ConvertToRange while catching conversion errors, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the existing file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of tables (ListObjects) on the current worksheet
                    ListObjectCollection tables = sheet.ListObjects;

                    // Iterate backwards because converting a table to a range removes it from the collection
                    for (int i = tables.Count - 1; i >= 0; i--)
                    {
                        ListObject table = tables[i];

                        try
                        {
                            // Convert the table to a regular range using default options
                            table.ConvertToRange();
                        }
                        catch (Exception ex)
                        {
                            // Use DisplayName as the table identifier (Name property may not be available in some versions)
                            string tableId = !string.IsNullOrEmpty(table.DisplayName) ? table.DisplayName : $"Index_{i}";
                            Console.WriteLine($"Failed to convert table '{tableId}' on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
