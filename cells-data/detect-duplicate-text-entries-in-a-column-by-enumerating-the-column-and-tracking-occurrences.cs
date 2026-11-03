// Title: Identify and list duplicate text values in an Excel column using Aspose.Cells for .NET (C#) and write results to a new worksheet
// Common Searches: C# Aspose.Cells find duplicate strings in a specific column and export them | how to count repeated cell values in column A with Aspose.Cells .NET | Aspose.Cells create a worksheet of duplicate entries from an existing Excel file | detect case‑insensitive duplicate text in Excel using Aspose.Cells C# example | enumerate rows of a column and collect duplicates with Aspose.Cells library
// Tags: duplicate detection in Excel column using Aspose.Cells | enumerate column cells with Aspose.Cells C# | case‑insensitive string counting in Aspose.Cells | export duplicate entries to separate worksheet using Aspose.Cells | save modified workbook to new file Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, iterates through all rows of column A, counts case‑insensitive string values with a dictionary, captures values that appear more than once, adds a new worksheet named "Duplicates" containing each duplicate value and its occurrence count, and saves the updated workbook to a new file.
class DuplicateDetector
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (add password here if required)
            LoadOptions loadOptions = new LoadOptions();
            // loadOptions.Password = "your_password"; // uncomment and set if needed
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the last row with data (zero‑based)
            int maxRow = cells.MaxDataRow;

            // Count occurrences of each text value (case‑insensitive)
            Dictionary<string, int> occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            List<string> duplicates = new List<string>();

            // Scan column A (index 0)
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = cells[row, 0];
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue.Trim();
                    if (string.IsNullOrEmpty(text))
                        continue;

                    if (occurrences.ContainsKey(text))
                    {
                        occurrences[text]++;
                        if (occurrences[text] == 2) // first time a duplicate is detected
                            duplicates.Add(text);
                    }
                    else
                    {
                        occurrences[text] = 1;
                    }
                }
            }

            // Output duplicate entries to the console
            Console.WriteLine("Duplicate entries in column A:");
            foreach (string dup in duplicates)
            {
                Console.WriteLine($"{dup} (occurs {occurrences[dup]} times)");
            }

            // Create a new worksheet for duplicates
            int dupSheetIndex = workbook.Worksheets.Add();
            Worksheet dupSheet = workbook.Worksheets[dupSheetIndex];
            dupSheet.Name = "Duplicates";

            // Header row
            dupSheet.Cells[0, 0].PutValue("Value");
            dupSheet.Cells[0, 1].PutValue("Count");

            // Populate duplicate data
            int dupRow = 1;
            foreach (string dup in duplicates)
            {
                dupSheet.Cells[dupRow, 0].PutValue(dup);
                dupSheet.Cells[dupRow, 1].PutValue(occurrences[dup]);
                dupRow++;
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
