// Title: Trim trailing whitespace after line breaks in Excel cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove trailing spaces after line break in cell | Trim whitespace from all cells in an Excel workbook using Aspose.Cells .NET | C# code to clean extra spaces after newline characters in worksheet with Aspose.Cells | How to save an Excel file without trailing spaces using Aspose.Cells | Process every worksheet to trim cell strings in Aspose.Cells C#
// Tags: trim trailing whitespace Aspose.Cells | remove extra spaces after newline C# | process all worksheets cell string trimming Aspose.Cells | save workbook without trailing spaces Aspose.Cells | cell string cleanup Excel .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing or new workbook, iterates through each worksheet and cell, trims trailing whitespace from string values, updates modified cells, ensures the output directory exists, and saves the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

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

                // Trim trailing spaces in all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        Cells cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = cells[row, col];
                                if (cell.Type == CellValueType.IsString)
                                {
                                    string original = cell.StringValue;
                                    if (!string.IsNullOrEmpty(original))
                                    {
                                        string trimmed = original.TrimEnd();
                                        if (trimmed != original)
                                        {
                                            cell.PutValue(trimmed);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with trimmed spaces
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
