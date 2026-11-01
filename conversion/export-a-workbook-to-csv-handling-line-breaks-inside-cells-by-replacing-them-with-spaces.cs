// Title: Export an Excel workbook to CSV in C# with Aspose.Cells while stripping line‑break characters from cell text
// Common Searches: c# aspnet replace line breaks in Excel cells before saving as csv using Aspose.Cells | how to remove newline characters from Excel cell values during CSV export with Aspose.Cells | Aspose.Cells export to CSV without embedded carriage returns in cell strings | clean Excel worksheet text for CSV output in .NET Aspose.Cells | remove \r\n from cells when converting XLSX to CSV with Aspose.Cells
// Tags: Aspose.Cells CSV export line break handling | C# clean Excel cell strings before CSV conversion | remove carriage return from workbook cells Aspose | sanitize worksheet text for CSV output | save XLSX as CSV with whitespace replacement

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLSX workbook using Aspose.Cells, iterates through every worksheet and cell, replaces Windows (\r\n), Unix (\n) and old Mac (\r) line‑break characters in string cells with a space, updates the cells when changes occur, and then saves the workbook directly as a CSV file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only string cells (line breaks are relevant for text)
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;

                            // Replace Windows (\r\n), Unix (\n) and old Mac (\r) line breaks with a space
                            string cleaned = original.Replace("\r\n", " ")
                                                    .Replace("\n", " ")
                                                    .Replace("\r", " ");

                            // Update the cell only if a change was made
                            if (cleaned != original)
                            {
                                cell.PutValue(cleaned);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook directly as CSV using the built‑in format enum
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
