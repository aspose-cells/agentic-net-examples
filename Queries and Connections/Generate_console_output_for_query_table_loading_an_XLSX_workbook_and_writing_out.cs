using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Path to the output text file
        string txtPath = "querytables.txt";

        // Load the workbook from the specified file (lifecycle: load)
        Workbook workbook = new Workbook(excelPath);

        // Create a TextWriter (StreamWriter) to write the query table information to a TXT file
        using (StreamWriter writer = new StreamWriter(txtPath))
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                writer.WriteLine($"Worksheet: {sheet.Name}");

                // Check if the worksheet contains any query tables
                if (sheet.QueryTables.Count > 0)
                {
                    // Iterate through each query table in the worksheet
                    foreach (QueryTable qt in sheet.QueryTables)
                    {
                        writer.WriteLine($"  Query Table Name: {qt.Name}");
                        writer.WriteLine($"  Result Range: {qt.ResultRange.Address}");
                        writer.WriteLine($"  Preserve Formatting: {qt.PreserveFormatting}");
                        writer.WriteLine($"  Adjust Column Width: {qt.AdjustColumnWidth}");
                    }
                }
                else
                {
                    writer.WriteLine("  No query tables found.");
                }

                writer.WriteLine(); // Blank line for readability
            }
        }

        // Output the generated TXT content to the console
        Console.WriteLine(File.ReadAllText(txtPath));
    }
}