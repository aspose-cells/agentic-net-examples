// Title: C# LightCellsDataHandler – Stream Excel rows to CSV with Aspose.Cells
// Description: Demonstrates a custom CsvExportHandler that inherits LightCellsDataHandler to read an Excel workbook in LightCells mode, collect each cell's string value, escape CSV delimiters, and write a comma‑separated line for every processed row. The handler is attached via LoadOptions, enabling memory‑efficient export of large workbooks to a CSV file.
// Keywords: Aspose.Cells LightCells CSV export | C# LightCellsDataHandler example | stream Excel to CSV .NET | memory efficient Excel export | large workbook CSV conversion | Aspose.Cells CSV writer | LightCells mode C# | Excel row streaming Aspose
// Common Searches: Aspose.Cells LightCells export to CSV C# | How to write Excel rows to CSV using LightCellsDataHandler | C# stream large Excel file to CSV with Aspose | LightCellsDataHandler CSV example | Memory‑efficient CSV export from Excel .NET
// Developer Intent: Export an Excel workbook to CSV by processing rows with LightCellsDataHandler for low‑memory usage.
// Use Cases: Generate CSV reports from massive Excel files without loading the entire workbook into memory. | Feed Excel data into external ETL pipelines that require CSV input. | Provide a lightweight CSV download endpoint in a web API using Aspose.Cells LightCells mode.
// AI Prompts: Show a C# Aspose.Cells LightCellsDataHandler that writes each processed row to a CSV file, handling commas, quotes, and line breaks. | Modify the CsvExportHandler to use a semicolon delimiter and skip rows where all cells are empty. | Explain how to integrate the LightCells CSV export into an ASP.NET Core controller that returns the generated CSV as a file download.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace LightCellsCsvExport
{
    // Custom LightCellsDataHandler that writes each processed row to a CSV file.
    // Demonstrates a custom CsvExportHandler that inherits LightCellsDataHandler to read an Excel workbook in LightCells mode, collect each cell's string value, escape CSV delimiters, and write a comma‑separated line for every processed row. The handler is attached via LoadOptions, enabling memory‑efficient export of large workbooks to a CSV file.
    class CsvExportHandler : LightCellsDataHandler
    {
        private readonly StreamWriter _writer;
        private readonly List<string> _currentRowValues = new List<string>();
        private bool _firstRow = true;

        public CsvExportHandler(string csvFilePath)
        {
            _writer = new StreamWriter(csvFilePath);
        }

        // Called when a worksheet starts processing.
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets.
            return true;
        }

        // Called before processing a row.
        public bool StartRow(int rowIndex)
        {
            // If this is not the first row, write the previously collected values.
            if (!_firstRow)
            {
                _writer.WriteLine(string.Join(",", _currentRowValues));
                _currentRowValues.Clear();
            }
            else
            {
                _firstRow = false;
            }

            // Continue processing this row.
            return true;
        }

        // Called after row properties are read (not used here, but required by interface).
        public bool ProcessRow(Row row)
        {
            return true; // Continue processing cells of this row.
        }

        // Determines whether a cell should be processed.
        public bool StartCell(int columnIndex)
        {
            // Process all cells.
            return true;
        }

        // Called for each cell; collect its string representation.
        public bool ProcessCell(Cell cell)
        {
            // Get the cell value as string, handling nulls.
            string value = cell.StringValue ?? string.Empty;

            // Escape CSV special characters.
            if (value.Contains("\""))
                value = value.Replace("\"", "\"\"");

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                value = $"\"{value}\"";

            _currentRowValues.Add(value);
            return true; // Keep processing.
        }

        // Flush remaining data and close the writer.
        public void Close()
        {
            if (_currentRowValues.Count > 0)
                _writer.WriteLine(string.Join(",", _currentRowValues));

            _writer.Flush();
            _writer.Dispose();
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths for the temporary Excel file and the resulting CSV.
            const string excelPath = "sample.xlsx";
            const string csvPath = "output.csv";

            // -------------------------------------------------
            // 1. Create a workbook and populate it with sample data.
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row.
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            // Sample data rows.
            for (int i = 0; i < 5; i++)
            {
                cells[i + 1, 0].PutValue(i + 1);                     // ID
                cells[i + 1, 1].PutValue($"Student_{i + 1}");       // Name
                cells[i + 1, 2].PutValue(60 + i * 5);               // Score
            }

            // Save the workbook normally (required for LightCells reading).
            workbook.Save(excelPath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // 2. Set up LightCellsDataHandler to export rows to CSV.
            // -------------------------------------------------
            var csvHandler = new CsvExportHandler(csvPath);
            var loadOptions = new LoadOptions
            {
                LightCellsDataHandler = csvHandler
            };

            // Load the workbook using LightCells mode; rows are streamed to the handler.
            Workbook loadedWorkbook = new Workbook(excelPath, loadOptions);

            // After loading, finalize CSV writing.
            csvHandler.Close();

            Console.WriteLine($"CSV export completed. File saved at: {Path.GetFullPath(csvPath)}");
        }
    }
}
