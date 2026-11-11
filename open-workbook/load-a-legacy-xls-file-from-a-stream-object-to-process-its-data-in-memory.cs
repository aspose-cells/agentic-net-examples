// Title: Load a legacy .xls workbook from a Stream and iterate its used cells using Aspose.Cells for .NET
// Common Searches: c# load .xls workbook from a memory stream with Aspose.Cells | read used range of an Excel file from Stream using Aspose.Cells .NET | process legacy Excel file without saving to disk Aspose.Cells C# | how to iterate non‑empty cells in a workbook loaded from Stream Aspose.Cells | Aspose.Cells open Excel file from FileStream and enumerate cells
// Tags: load legacy xls stream Aspose.Cells | auto detect workbook format Aspose.Cells | enumerate used range cells Aspose.Cells | read Excel from memory stream C# | access first worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates loading a legacy .xls workbook directly from a Stream (or FileStream) with Aspose.Cells for .NET, automatically detecting the format, accessing the first worksheet, determining the used range, and iterating over non‑empty cells to output their values, while handling null streams and exceptions.
    public class ExcelProcessor
    {
        /// <param name="xlsStream">Stream containing the Excel file.</param>
        public void LoadAndProcess(Stream xlsStream)
        {
            if (xlsStream == null) throw new ArgumentNullException(nameof(xlsStream));

            try
            {
                // Ensure the stream is positioned at the beginning
                if (xlsStream.CanSeek)
                    xlsStream.Position = 0;

                // Load the workbook (format is auto‑detected)
                Workbook workbook = new Workbook(xlsStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Iterate through used cells
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Skip empty cells
                        if (cell.Type != CellValueType.IsNull)
                        {
                            Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect a file path as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsExample <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    ExcelProcessor processor = new ExcelProcessor();
                    processor.LoadAndProcess(fs);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
