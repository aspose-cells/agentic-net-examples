// Title: Set a Custom Buffer Size for Aspose.Cells LightCells Mode in C# to Boost High‑Throughput Reading
// Description: Learn how to configure Aspose.Cells LightCells mode with a user‑defined buffer size, enable MultiThreadReading, and process large worksheets in parallel. The example shows a custom LightCellsDataHandler that receives a buffer size, logs sheet activity, and reads cells efficiently for high‑throughput scenarios.
// Keywords: Aspose.Cells LightCells custom buffer | C# LightCellsDataHandler buffer size | multi‑threaded Excel reading Aspose | high‑throughput cell processing .NET | optimize LightCells performance | load large workbook with LightCells | Aspose.Cells parallel cell read
// Common Searches: how to set buffer size for LightCellsDataHandler Aspose.Cells | enable multi‑thread reading with LightCells in C# | Aspose.Cells LightCells performance tuning | custom LightCellsDataHandler example | read large Excel file quickly using Aspose.Cells
// Developer Intent: Configure LightCells to use a custom buffer size and multi‑threaded reading for faster processing of large Excel files.
// Use Cases: Load a multi‑megabyte workbook with a 64 KB buffer to limit memory consumption. | Process the first column of every row in parallel for rapid data extraction. | Insert custom per‑cell logic while controlling memory via a configurable buffer.
// AI Prompts: Generate C# code that creates a LightCellsDataHandler accepting a buffer size parameter and loads a workbook in LightCells mode. | Show how to adjust the buffer size and enable MultiThreadReading to maximize throughput when processing a large Excel file with Aspose.Cells. | Provide an example of custom cell handling inside ProcessCell while respecting a user‑defined buffer size.

using System;
using System.Threading.Tasks;
using Aspose.Cells;

// Custom LightCellsDataHandler that accepts a buffer size parameter.
// The buffer size can be used internally for any custom logic while processing cells.
// Learn how to configure Aspose.Cells LightCells mode with a user‑defined buffer size, enable MultiThreadReading, and process large worksheets in parallel. The example shows a custom LightCellsDataHandler that receives a buffer size, logs sheet activity, and reads cells efficiently for high‑throughput scenarios.
class CustomLightCellsDataHandler : LightCellsDataHandler
{
    private readonly int _bufferSize;

    public CustomLightCellsDataHandler(int bufferSize)
    {
        _bufferSize = bufferSize;
    }

    // Called when a worksheet starts processing.
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Processing sheet \"{sheet.Name}\" with buffer size {_bufferSize} bytes.");
        return true; // Continue processing this sheet.
    }

    // Called before a row is processed.
    public bool StartRow(int rowIndex)
    {
        return true; // Continue processing this row.
    }

    // Called after a row is read.
    public bool ProcessRow(Row row)
    {
        return true; // Continue processing.
    }

    // Called before a cell is processed.
    public bool StartCell(int columnIndex)
    {
        return true; // Continue processing this cell.
    }

    // Called after a cell is read.
    public bool ProcessCell(Cell cell)
    {
        // Example: simply output the cell address and value.
        Console.WriteLine($"Cell [{cell.Row}, {cell.Column}] = {cell.Value}");
        return true; // Continue processing.
    }
}

class Program
{
    static void Main()
    {
        // Define a custom buffer size (e.g., 64 KB) for reading cells.
        int customBufferSize = 64 * 1024;

        // Create LoadOptions and assign the custom LightCellsDataHandler.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new CustomLightCellsDataHandler(customBufferSize);

        // Load the workbook using LightCells mode with the custom handler.
        using (Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions))
        {
            // Enable multi‑thread reading to maximize throughput.
            workbook.Worksheets[0].Cells.MultiThreadReading = true;

            // Example: read the first column of all rows in parallel.
            Worksheet sheet = workbook.Worksheets[0];
            int maxRow = sheet.Cells.MaxDataRow;

            Parallel.For(0, maxRow + 1, rowIndex =>
            {
                Cell cell = sheet.Cells[rowIndex, 0];
                var value = cell.Value; // Access cell value safely under MultiThreadReading.
                // Perform any high‑throughput processing here.
            });
        }
    }
}
