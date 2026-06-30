using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Author: Aspose.Cells LightCells CSV exporter
class Program
{
    static void Main()
    {
        // Input workbook (can be any supported format)
        const string inputPath = "input.xlsx";
        // Output CSV file where each processed row will be written
        const string outputCsvPath = "output.csv";

        // Prepare a StreamWriter for the CSV output
        using (var csvWriter = new StreamWriter(outputCsvPath, false, Encoding.UTF8))
        {
            // Create custom LightCellsDataHandler that writes rows to CSV
            var handler = new CsvExportLightCellsDataHandler(csvWriter);

            // Set up load options to use the custom handler
            var loadOptions = new LoadOptions
            {
                LightCellsDataHandler = handler
            };

            // Load the workbook in LightCells mode; processing occurs via the handler
            using (var workbook = new Workbook(inputPath, loadOptions))
            {
                // No further action required – rows are written during loading
            }
        }

        Console.WriteLine("CSV export completed.");
    }
}

// Custom LightCellsDataHandler that streams each row to a CSV file
class CsvExportLightCellsDataHandler : LightCellsDataHandler
{
    private readonly StreamWriter _writer;
    private readonly StringBuilder _rowBuilder = new StringBuilder();

    public CsvExportLightCellsDataHandler(StreamWriter writer)
    {
        _writer = writer;
    }

    // Process all worksheets
    public bool StartSheet(Worksheet sheet)
    {
        // Optionally write a sheet header (e.g., sheet name) – omitted for plain CSV
        return true; // Continue processing this sheet
    }

    // Called before processing a row; reset the row buffer
    public bool StartRow(int rowIndex)
    {
        _rowBuilder.Clear();
        return true; // Continue processing this row
    }

    // Called before processing a cell; nothing special needed here
    public bool StartCell(int columnIndex)
    {
        return true; // Continue processing this cell
    }

    // Process each cell: append its value to the row buffer, escaping as needed
    public bool ProcessCell(Cell cell)
    {
        // Retrieve the cell's string representation
        string cellText = cell.Value?.ToString() ?? string.Empty;

        // Escape double quotes by doubling them
        if (cellText.Contains("\""))
            cellText = cellText.Replace("\"", "\"\"");

        // If the cell contains a comma, newline or quote, wrap it in quotes
        if (cellText.Contains(",") || cellText.Contains("\n") || cellText.Contains("\""))
            cellText = $"\"{cellText}\"";

        // Append value and a comma separator
        if (_rowBuilder.Length > 0)
            _rowBuilder.Append(',');

        _rowBuilder.Append(cellText);
        return true; // Continue processing
    }

    // Called after all cells of the row have been processed; write the row to CSV
    public bool ProcessRow(Row row)
    {
        _writer.WriteLine(_rowBuilder.ToString());
        return true; // Continue processing subsequent rows
    }
}