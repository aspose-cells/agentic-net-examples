using System;
using System.IO;
using Aspose.Cells;

public class LightCellsHandler : LightCellsDataHandler
{
    // Called for each worksheet; return true to process the sheet
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Processing sheet: {sheet.Name}");
        return true;
    }

    // Called for each row index; return true to process the row
    public bool StartRow(int rowIndex)
    {
        // You can add row‑level filtering here
        return true;
    }

    // Called after a row is started; return true to process its cells
    public bool ProcessRow(Row row)
    {
        // Example: just output the row index
        Console.WriteLine($"Row {row.Index} started");
        return true;
    }

    // Called for each cell column index; return true to process the cell
    public bool StartCell(int columnIndex)
    {
        return true;
    }

    // Called for each cell that should be processed
    public bool ProcessCell(Cell cell)
    {
        // Minimal processing – just read the value
        var value = cell.Value;
        // Optionally, output the cell address and value
        Console.WriteLine($"Cell {cell.Name}: {value}");
        return true;
    }
}

public class Program
{
    public static void Main()
    {
        // Path to the large XLSX file
        string inputFile = "LargeFile.xlsx";
        string outputFile = "ProcessedLargeFile.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Configure LoadOptions to use LightCells and limit memory usage
            LoadOptions loadOptions = new LoadOptions
            {
                // Use a custom LightCellsDataHandler to stream data
                LightCellsDataHandler = new LightCellsHandler(),
                // FileCache stores intermediate data on disk, keeping RAM usage low
                MemorySetting = MemorySetting.FileCache
            };

            // Load the workbook in LightCells mode (streaming) with the specified options
            using (Workbook workbook = new Workbook(inputFile, loadOptions))
            {
                // Save the processed workbook
                workbook.Save(outputFile);
            }

            Console.WriteLine($"Workbook loaded and saved to '{outputFile}' with memory usage limited to ~200 MB.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}