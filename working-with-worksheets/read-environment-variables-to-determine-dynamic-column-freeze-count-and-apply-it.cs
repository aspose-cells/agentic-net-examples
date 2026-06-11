using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDynamicFreeze
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default single worksheet)
                Workbook workbook = new Workbook();

                // Read the environment variable that specifies how many columns to freeze
                // Example: set FREEZE_COLUMNS=3 in the environment before running
                string envValue = Environment.GetEnvironmentVariable("FREEZE_COLUMNS") ?? string.Empty;
                int freezeColumns = 0;

                // Try to parse the value; if parsing fails or is negative, default to 0 (no freeze)
                if (!int.TryParse(envValue, out freezeColumns) || freezeColumns < 0)
                {
                    freezeColumns = 0;
                }

                // Apply the freeze setting to each worksheet only when freezeColumns > 0
                if (freezeColumns > 0)
                {
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Freeze only columns. Row index = 0 (top row), column index = freezeColumns.
                        // Use the 4‑parameter overload required by the current Aspose.Cells version.
                        sheet.FreezePanes(0, freezeColumns, 0, 0);
                    }
                }

                // Define output file path
                string outputPath = "DynamicFreezeColumns.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}