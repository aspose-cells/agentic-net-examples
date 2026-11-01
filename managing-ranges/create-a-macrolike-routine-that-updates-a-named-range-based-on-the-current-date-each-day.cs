// Title: Update an Excel named range with today's date using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set today's date in a named range | Update Excel named range daily with Aspose.Cells .NET | Aspose.Cells replace values in a named range with current date | Programmatically refresh a named range with DateTime.Today using Aspose.Cells | How to apply date format to cells in a named range with Aspose.Cells C#
// Tags: Aspose.Cells update named range cells | set date value in Excel named range C# | apply built-in date number format Aspose.Cells | load workbook modify range and save Aspose.Cells | handle errors when updating named ranges Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// Provides a C# class that loads an Excel workbook with Aspose.Cells, retrieves a specified named range, writes the current date (DateTime.Today) into every cell of the range, applies the built‑in date number format, and saves the file, including argument validation and robust exception handling.
public class NamedRangeUpdater
{
    /// <param name="workbookPath">Full path to the Excel file.</param>
    /// <param name="namedRange">The name of the range to update.</param>
    public void UpdateNamedRange(string workbookPath, string namedRange)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            throw new ArgumentException("Workbook path cannot be null or empty.", nameof(workbookPath));

        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("The specified workbook file was not found.", workbookPath);

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the named range as a Range object
            AsposeRange range = workbook.Worksheets.GetRangeByName(namedRange);
            if (range == null)
                throw new ArgumentException($"Named range '{namedRange}' does not exist in the workbook.", nameof(namedRange));

            Worksheet sheet = range.Worksheet;

            int startRow = range.FirstRow;
            int endRow = startRow + range.RowCount - 1;
            int startCol = range.FirstColumn;
            int endCol = startCol + range.ColumnCount - 1;

            // Iterate through all cells in the range and set the current date
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue(DateTime.Today);

                    // Apply a built‑in date format (e.g., mm-dd-yy)
                    Style style = cell.GetStyle();
                    style.Number = 14;
                    cell.SetStyle(style);
                }
            }

            // Save the workbook (overwrites the original file)
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace while providing context
            throw new InvalidOperationException("Failed to update the named range.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage
            string workbookPath = @"C:\Reports\DailyReport.xlsx";
            string namedRange = "ReportDate";

            var updater = new NamedRangeUpdater();
            updater.UpdateNamedRange(workbookPath, namedRange);

            Console.WriteLine("Named range updated successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
