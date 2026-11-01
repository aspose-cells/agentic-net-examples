// Title: Export only visible rows from an Excel workbook to a tab‑delimited TXT file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through rows, skips any row where Cells.Rows[row].IsHidden is true, and writes the remaining cells to a tab‑separated .txt file. | Show how to add robust file‑existence checking and exception handling when exporting visible worksheet rows to a text file with Aspose.Cells. | Demonstrate using StreamWriter to create a tab‑delimited text export of visible data from the first worksheet of a workbook loaded by Aspose.Cells.
// Common Searches: C# Aspose.Cells export visible rows to tab delimited text file | How to ignore hidden rows when saving Excel data as .txt using Aspose.Cells | Export first worksheet data without hidden rows to TXT in .NET | Aspose.Cells write visible cells to text file with tab separators | Skip hidden rows during Excel to TXT conversion in C#
// Tags: Aspose.Cells visible rows text export | filter hidden rows Aspose.Cells C# | tab delimited Excel to TXT conversion | worksheet row visibility check Aspose.Cells | streamwriter tab separated output .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, iterates over the first worksheet, skips rows flagged as hidden, and writes the visible cell values to a tab‑separated text file using StreamWriter.
class ExportVisibleRowsToTxt
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a StreamWriter for the output TXT file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Determine the range of used rows and columns
                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
                {
                    // Skip hidden rows using the Row object from Cells.Rows collection
                    if (sheet.Cells.Rows[rowIndex].IsHidden)
                        continue;

                    // Collect cell values for the current visible row
                    string[] values = new string[maxColumn + 1];
                    for (int colIndex = 0; colIndex <= maxColumn; colIndex++)
                    {
                        Cell cell = sheet.Cells[rowIndex, colIndex];
                        values[colIndex] = cell?.StringValue ?? string.Empty;
                    }

                    // Write the row to the TXT file, separating values with a tab character
                    writer.WriteLine(string.Join("\t", values));
                }
            }

            Console.WriteLine($"Visible data exported to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during export:");
            Console.WriteLine(ex.Message);
        }
    }
}
