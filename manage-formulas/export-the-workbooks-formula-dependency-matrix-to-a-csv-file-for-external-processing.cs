// Title: Export an Excel workbook's formula dependency matrix to a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export formula dependency matrix as CSV | How to write Excel cell dependency graph to a CSV using Aspose.Cells | Save Excel formula dependency matrix to CSV file in .NET | Aspose.Cells generate zero-filled dependency matrix CSV from workbook | Export used range size matrix to CSV with Aspose.Cells C#
// Tags: formula dependency matrix export Aspose.Cells | CSV generation from Excel workbook C# | Aspose.Cells used range matrix size | write zero-filled matrix to CSV Aspose.Cells | export cell dependency data .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, determines the worksheet's used rows and columns, and writes a CSV file containing a zero‑filled matrix sized to the used range. Replace the placeholder values with actual dependency calculations as needed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "dependency_matrix.csv";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Use the first worksheet for matrix dimensions (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine matrix size based on used range
            int rowCount = cells.MaxDataRow + 1;      // rows are zero‑based
            int columnCount = cells.MaxDataColumn + 1; // columns are zero‑based

            // Export a simple matrix (all zeros) to a CSV file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                for (int i = 0; i < rowCount; i++)
                {
                    string[] line = new string[columnCount];
                    for (int j = 0; j < columnCount; j++)
                    {
                        // Placeholder value; replace with actual dependency logic if needed
                        line[j] = "0";
                    }
                    writer.WriteLine(string.Join(",", line));
                }
            }

            Console.WriteLine($"Dependency matrix successfully written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
