// Title: Set PreserveFormatting = true for all QueryTables in an Aspose.Cells workbook (C# .NET)
// Description: Loads a workbook (or creates a new one), loops through every worksheet and each QueryTable, enables the PreserveFormatting flag, and saves the updated file.
// Keywords: Aspose.Cells PreserveFormatting | QueryTable PreserveFormatting C# | set query table formatting Aspose.Cells | iterate query tables .NET | Excel external data connection formatting | C# Aspose.Cells QueryTable property | preserve cell style after data refresh
// Common Searches: how to enable PreserveFormatting for all QueryTables in Aspose.Cells | C# loop through worksheets and set QueryTable PreserveFormatting | Aspose.Cells keep formatting of external data connections | set PreserveFormatting property for QueryTables using .NET | Aspose.Cells preserve cell style after refresh
// Developer Intent: Enable the PreserveFormatting flag on every QueryTable across all worksheets in a workbook.
// Use Cases: Ensure custom cell styles remain after refreshing external data sources. | Prepare workbooks for distribution where formatting must stay consistent regardless of data changes. | Automate batch processing to enforce uniform formatting for all QueryTables before publishing.
// AI Prompts: Generate C# code with Aspose.Cells that sets PreserveFormatting = true for every QueryTable in a workbook and saves it. | Create a reusable method that accepts a Workbook object and activates PreserveFormatting on all its QueryTables. | Provide robust error handling for loading a workbook, updating QueryTable properties, and saving the file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook (or creates a new one), loops through every worksheet and each QueryTable, enables the PreserveFormatting flag, and saves the updated file.
    public class SetQueryTablePreserveFormatting
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Enable preserving formatting for all query tables in all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (QueryTable queryTable in worksheet.QueryTables)
                    {
                        queryTable.PreserveFormatting = true;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetQueryTablePreserveFormatting.Run();
        }
    }
}
