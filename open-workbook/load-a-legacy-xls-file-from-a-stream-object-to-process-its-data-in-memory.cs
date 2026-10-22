// Title: Load a legacy XLS workbook from a Stream and iterate its cells using Aspose.Cells for .NET
// Common Searches: how to open an old .xls file from a MemoryStream with Aspose.Cells in C# | reading cell values from a legacy Excel workbook without saving to disk using Aspose.Cells | Aspose.Cells LoadOptions auto-detect XLS format from a stream | enumerate used range of a workbook loaded from a stream in Aspose.Cells .NET | process in‑memory Excel 97‑2003 file with Aspose.Cells C#
// Tags: load xls workbook from stream Aspose.Cells | auto-detect xls format with LoadOptions | enumerate used cells Aspose.Cells | read first worksheet cell values C# | in-memory legacy Excel processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates loading a legacy XLS workbook directly from a Stream using Aspose.Cells' Workbook constructor with LoadOptions, accessing the first worksheet, determining the used range via MaxDataRow/MaxDataColumn, iterating each cell to output its address and value, and handling errors with try‑catch.
    public class ExcelProcessor
    {
        public void ProcessXls(Stream xlsStream)
        {
            try
            {
                // Load the workbook from the provided stream (auto-detect format)
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(xlsStream, loadOptions);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Iterate through the used cells and output values
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        object value = cells[row, col].Value;
                        Console.WriteLine($"R{row + 1}C{col + 1}: {value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing XLS stream: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string filePath = args.Length > 0 ? args[0] : "sample.xls";

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                using (FileStream stream = File.OpenRead(filePath))
                {
                    ExcelProcessor processor = new ExcelProcessor();
                    processor.ProcessXls(stream);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
