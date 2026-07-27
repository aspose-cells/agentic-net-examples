// Title: Insert Hyperlinks While Streaming Data Using LightCellsDataProvider in Aspose.Cells for .NET
// Description: Demonstrates creating a Workbook, adding an external hyperlink to cell B1, and populating the first worksheet via a custom LightCellsDataProvider. The provider streams rows in memory‑efficient mode, and the pre‑added hyperlink is preserved when saving the file with OoxmlSaveOptions (XLSX).
// Keywords: Aspose.Cells | LightCellsDataProvider | C# hyperlink Excel | streaming Excel data | memory efficient workbook | OoxmlSaveOptions | add hyperlink programmatically | large Excel export | dynamic URL links | GitHub Aspose.Cells example
// Common Searches: how to add a hyperlink when using LightCellsDataProvider Aspose.Cells | stream rows to Excel and keep hyperlinks .NET | Aspose.Cells LightCellsDataProvider example with links | preserve hyperlinks in light mode save | C# insert hyperlink before streaming data to workbook
// Developer Intent: Write Excel data with LightCellsDataProvider while embedding hyperlinks that remain functional after saving.
// Use Cases: Generate a massive report where each row contains a link to a detailed web page, using LightCellsDataProvider to stream rows without loading the whole workbook into memory. | Create a template with a static navigation link (e.g., B1) and then fill additional rows with data, ensuring the link stays active in the final XLSX file. | Export millions of records to an Excel file in a memory‑constrained environment while embedding external URLs for quick navigation. | Build an automated data pipeline that streams CSV data into Excel and adds contextual hyperlinks on the fly.
// AI Prompts: Show how to modify CustomLightCellsDataProvider so that each cell in column B receives a hyperlink based on its row ID. | Provide C# code that sets the hyperlink display text after writing the cell value when using LightCellsDataProvider. | Explain the steps to configure OoxmlSaveOptions with LightCellsDataProvider to retain all hyperlinks in the saved workbook. | Generate a GitHub‑ready README snippet that describes this example and includes a badge for Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates creating a Workbook, adding an external hyperlink to cell B1, and populating the first worksheet via a custom LightCellsDataProvider. The provider streams rows in memory‑efficient mode, and the pre‑added hyperlink is preserved when saving the file with OoxmlSaveOptions (XLSX).
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a hyperlink to cell B1 (will be filled later by the data provider)
        // The hyperlink points to an external URL and the display text will be set after the cell value is written
        sheet.Hyperlinks.Add("B1", 1, 1, "https://www.example.com");
        sheet.Cells["B1"].PutValue("Visit Example");

        // Create a custom LightCellsDataProvider that supplies cell data in streaming mode
        var dataProvider = new CustomLightCellsDataProvider();

        // Configure save options to use the LightCellsDataProvider
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = dataProvider
        };

        // Save the workbook using the light mode options
        workbook.Save("HyperlinkLightCells.xlsx", saveOptions);
    }

    // Custom implementation of LightCellsDataProvider
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        // Sample data to be written to the worksheet
        // First column: ID, second column: placeholder for hyperlink text
        private readonly string[,] _data = new string[,]
        {
            { "ID", "Link" },
            { "1", "" },
            { "2", "" },
            { "3", "" }
        };

        private int _currentRow = -1;
        private int _currentColumn = -1;

        // Process only the first worksheet (index 0)
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index to be saved; -1 signals no more rows
        public int NextRow()
        {
            _currentRow++;
            _currentColumn = -1;
            return _currentRow < _data.GetLength(0) ? _currentRow : -1;
        }

        // Called before writing cells of the current row; no special handling needed here
        public void StartRow(Row row) { }

        // Return the next column index within the current row; -1 signals end of columns for the row
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn < _data.GetLength(1) ? _currentColumn : -1;
        }

        // Populate the current cell with the predefined data
        public void StartCell(Cell cell)
        {
            cell.PutValue(_data[_currentRow, _currentColumn]);
        }

        // Indicate that string values should be gathered into a global string pool for efficiency
        public bool IsGatherString() => true;
    }
}
