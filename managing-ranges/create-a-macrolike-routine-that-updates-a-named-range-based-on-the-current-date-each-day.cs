// Title: C# method to write the current date into a specific named range in an Excel workbook with Aspose.Cells
// AI Prompts: Write C# code that opens an existing .xlsx file using Aspose.Cells, finds a named range, inserts DateTime.Today into its top‑left cell, applies the short date number format, and saves the workbook. | Create a static helper function that takes a workbook path and a named range name, updates that range with today's date, formats the cell as m/d/yyyy, and includes error handling for missing files or undefined ranges.
// Common Searches: asp.net cells how to set today's date in a named range of an existing workbook | c# update excel named range with current date using Aspose.Cells library | programmatically refresh a date cell in a named range each day with Aspose.Cells | Aspose.Cells C# example for writing DateTime.Today to the first cell of a named range
// Tags: update named range with current date Aspose.Cells | write DateTime.Today to first cell of named range | apply short date number format cell Aspose.Cells | load and save workbook after modifying named range | error handling missing file or undefined named range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The provided C# example defines a static UpdateNamedRange method that loads an Excel workbook, retrieves a specified named range, writes the current date (without time) into its first cell, applies the m/d/yyyy number format, and saves the file, with comprehensive error handling for missing files and undefined ranges.
public class NamedRangeUpdater
{
    // Updates the specified named range with the current date (date only, no time).
    public static void UpdateNamedRange(string workbookPath, string namedRange)
    {
        try
        {
            // Ensure the workbook file exists before attempting to load it.
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            // Load the existing workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the range associated with the named range.
            AsposeRange range = workbook.Worksheets.GetRangeByName(namedRange);
            if (range == null)
                throw new ArgumentException($"Named range '{namedRange}' does not exist in the workbook.");

            // Determine the top‑left cell of the range.
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            Worksheet sheet = range.Worksheet;

            // Write the current date (without time) into the cell.
            Cell targetCell = sheet.Cells[firstRow, firstColumn];
            targetCell.PutValue(DateTime.Today);

            // Apply a date number format (m/d/yyyy).
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // 14 = "m/d/yyyy"
            targetCell.SetStyle(dateStyle);

            // Save the workbook (overwrites the original file).
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error updating named range: {ex.Message}");
            throw;
        }
    }

    // Example usage.
    public static void Main()
    {
        try
        {
            string filePath = @"C:\Temp\Sample.xlsx";
            string rangeName = "TodayDate";

            UpdateNamedRange(filePath, rangeName);
            Console.WriteLine("Named range updated with current date.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
