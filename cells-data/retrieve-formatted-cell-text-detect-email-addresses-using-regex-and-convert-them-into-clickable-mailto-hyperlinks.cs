// Title: Detect email addresses in Excel cells and convert them to mailto hyperlinks using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add mailto hyperlink to cells containing email addresses | How to use regex with Aspose.Cells to find and hyperlink emails in an .xlsx workbook | Iterate over used range in Aspose.Cells and create clickable email links in C# | Replace email text with mailto link programmatically using Aspose.Cells
// Tags: regex email detection Aspose.Cells | email hyperlink insertion Aspose.Cells | iterate used range cells Aspose.Cells C# | programmatic email link creation Excel .xlsx | hyperlink API worksheet Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an input.xlsx workbook with Aspose.Cells, scans each worksheet's used range, uses a regular expression to locate email addresses in cell text, adds a mailto hyperlink that covers the entire cell for each match, and saves the modified workbook to output.xlsx while handling file‑not‑found and loading errors.
class EmailHyperlinkConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password required). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            LoadOptions loadOptions = new LoadOptions();
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        // Regular expression pattern to detect email addresses
        string emailPattern = @"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}";

        try
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        string cellText = cell.StringValue;

                        if (string.IsNullOrEmpty(cellText))
                            continue;

                        // Search for an email address within the cell text
                        Match match = Regex.Match(cellText, emailPattern);
                        if (match.Success)
                        {
                            string email = match.Value;
                            // Add a mailto hyperlink to the cell (covers the whole cell)
                            sheet.Hyperlinks.Add(row, col, 1, 1, "mailto:" + email);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
