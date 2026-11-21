// Title: Update an Excel named range with the current date using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells code to set today's date in a named range of an existing workbook | How to programmatically fill a defined name with the current date each day using Aspose.Cells | Iterate over cells in a named range and apply a date number format with Aspose.Cells .NET | Overwrite an Excel file after updating a named range to the current date in C# | Retrieve and modify a named range by name with Aspose.Cells and save changes
// Tags: update named range cells Aspose.Cells | set current date Excel defined name C# | apply built‑in date format Aspose.Cells | iterate range cells Aspose.Cells API | save workbook after modifying named range Aspose

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace Example
{
    // The example loads an existing workbook, retrieves a named range, iterates through each cell in that range, writes the current date with a built‑in date number format, and saves the workbook, providing a reusable C# routine for daily date updates using Aspose.Cells.
    public class NamedRangeUpdater
    {
        /// <param name="workbookPath">Full path to the Excel file.</param>
        /// <param name="namedRange">The name of the range to update.</param>
        public void UpdateNamedRange(string workbookPath, string namedRange)
        {
            // Verify that the workbook file exists before attempting to load it
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(workbookPath);

                // Retrieve the range by its defined name
                AsposeRange range = workbook.Worksheets.GetRangeByName(namedRange);
                if (range == null)
                    throw new ArgumentException($"Named range '{namedRange}' does not exist in the workbook.");

                // Determine the bounds of the range
                int startRow = range.FirstRow;
                int endRow = range.FirstRow + range.RowCount - 1;
                int startColumn = range.FirstColumn;
                int endColumn = range.FirstColumn + range.ColumnCount - 1;
                Worksheet ws = range.Worksheet;

                // Current date (without time component)
                DateTime today = DateTime.Today;

                // Iterate through all cells in the range and set the value to today's date
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = ws.Cells[row, col];
                        cell.PutValue(today);

                        // Set the cell's style to display as a date
                        Style style = cell.GetStyle();
                        style.Number = 14; // Built‑in date format (e.g., mm-dd-yy)
                        cell.SetStyle(style);
                    }
                }

                // Save the workbook (overwrites the original file)
                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating named range: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <exe> <workbookPath> <namedRange>");
                return;
            }

            string workbookPath = args[0];
            string namedRange = args[1];

            try
            {
                var updater = new NamedRangeUpdater();
                updater.UpdateNamedRange(workbookPath, namedRange);
                Console.WriteLine("Named range updated successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Operation failed: {ex.Message}");
            }
        }
    }
}
