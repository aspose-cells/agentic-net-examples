// Title: How to iterate over mapped cell areas in an Excel worksheet and log row and column indices using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to enumerate mapped areas in a worksheet | Aspose.Cells .NET log row and column indices of XML mapped cells | Using reflection to get MappedAreas collection with Aspose.Cells | Debug Excel XML map cell ranges with Aspose.Cells C# example | Iterate used cells when MappedAreas property is missing in Aspose.Cells
// Tags: XML map region enumeration Aspose.Cells | log cell row column indices C# | reflection retrieve XML map areas | fallback used cells iteration Aspose.Cells | debug XML map ranges Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace MappedAreaDebugger
{
    // The sample loads an Excel workbook, selects the first worksheet, and attempts to obtain its MappedAreas collection via reflection. For each mapped area it logs the start and end row/column indices and then prints every cell coordinate inside the area. If the MappedAreas property is unavailable, the code falls back to iterating all used cells, logging their row, column, and value. All operations are wrapped in error handling to report issues.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Resolve workbook path and ensure the file exists.
                string workbookPath = GetWorkbookPath(args);
                Console.WriteLine($"Loading workbook: {workbookPath}");

                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Verify that at least one worksheet exists.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Use the first worksheet for demonstration.
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Processing worksheet: {sheet.Name}");

                // Try to obtain the MappedAreas collection via reflection (available in newer versions).
                var mappedAreasProp = sheet.GetType().GetProperty("MappedAreas");
                if (mappedAreasProp != null)
                {
                    var mappedAreas = mappedAreasProp.GetValue(sheet) as IEnumerable;
                    if (mappedAreas != null)
                    {
                        int areaCount = 0;
                        foreach (var area in mappedAreas)
                        {
                            areaCount++;
                            // Retrieve area boundaries using reflection.
                            int startRow = (int)area.GetType().GetProperty("StartRow").GetValue(area);
                            int startColumn = (int)area.GetType().GetProperty("StartColumn").GetValue(area);
                            int endRow = (int)area.GetType().GetProperty("EndRow").GetValue(area);
                            int endColumn = (int)area.GetType().GetProperty("EndColumn").GetValue(area);

                            Console.WriteLine(
                                $"Mapped Area {areaCount} - StartRow: {startRow}, StartColumn: {startColumn}, " +
                                $"EndRow: {endRow}, EndColumn: {endColumn}");

                            // Log each cell inside the mapped area.
                            for (int row = startRow; row <= endRow; row++)
                            {
                                for (int col = startColumn; col <= endColumn; col++)
                                {
                                    Console.WriteLine($"  Cell at Row={row}, Column={col}");
                                }
                            }
                        }

                        if (areaCount == 0)
                        {
                            Console.WriteLine("No mapped cell areas were found in this worksheet.");
                        }
                        else
                        {
                            Console.WriteLine($"Found {areaCount} mapped area(s).");
                        }

                        Console.WriteLine("Iteration over mapped areas completed successfully.");
                        return;
                    }
                }

                // Fallback: iterate over all used cells if MappedAreas are unavailable.
                Console.WriteLine("MappedAreas collection not available. Falling back to iterating used cells.");

                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell != null && cell.Type != CellValueType.IsNull)
                        {
                            Console.WriteLine($"Cell with data at Row={row}, Column={col}, Value={cell.StringValue}");
                        }
                    }
                }

                Console.WriteLine("Iteration over used cells completed successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="args">Command‑line arguments.</param>
        /// <returns>Full path to the Excel file.</returns>
        private static string GetWorkbookPath(string[] args)
        {
            const string defaultFileName = "Sample.xlsx";

            if (args != null && args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                string candidate = args[0];
                if (File.Exists(candidate))
                {
                    return Path.GetFullPath(candidate);
                }

                // Resolve relative path against the executable folder.
                string combined = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, candidate);
                if (File.Exists(combined))
                {
                    return Path.GetFullPath(combined);
                }

                throw new FileNotFoundException($"Specified workbook file not found: {candidate}");
            }

            // No argument supplied – look for the default file next to the executable.
            string defaultPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, defaultFileName);
            if (!File.Exists(defaultPath))
            {
                throw new FileNotFoundException($"Default workbook file not found: {defaultPath}");
            }

            return defaultPath;
        }
    }
}
