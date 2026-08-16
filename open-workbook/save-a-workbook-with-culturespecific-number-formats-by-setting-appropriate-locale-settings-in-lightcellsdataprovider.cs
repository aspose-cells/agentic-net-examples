// Title: Save Excel with locale‑specific number formats using a custom LightCellsDataProvider (Aspose.Cells for .NET)
// Description: Shows how to assign Workbook.Settings.CultureInfo inside a custom LightCellsDataProvider, copy cell values and styles, and export the workbook with OoxmlSaveOptions in LightCells mode so numbers follow the chosen locale (e.g., German separators).
// Keywords: Aspose.Cells | C# | LightCellsDataProvider | CultureInfo | locale number format | German decimal separator | Excel export options | OoxmlSaveOptions | culture‑aware Excel save | large workbook performance
// Common Searches: Aspose.Cells set workbook culture before saving | LightCellsDataProvider example C# | Save Excel with German number format using Aspose | Apply CultureInfo to Excel export .NET | Locale specific number formatting in Aspose.Cells
// Developer Intent: Apply a chosen CultureInfo during LightCells saving to generate locale‑aware numeric formatting.
// Use Cases: Produce financial statements for different regions where decimal and thousand separators vary, without loading the whole file into memory. | Export massive data sets with high performance while automatically applying the correct regional number format. | Create Excel reports for international users that display numbers according to their local conventions.
// AI Prompts: Extend the CultureAwareLightCellsDataProvider to support multiple worksheets, each with its own CultureInfo. | Explain how to combine Settings.CultureInfo with LightCellsDataProvider to format dates and currencies for a specific locale. | Provide a step‑by‑step guide to export a workbook to PDF while preserving locale‑specific number formatting using LightCells.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureSpecificSave
{
    // Custom LightCellsDataProvider that sets the workbook's culture before saving
    // and supplies cell data from the original workbook.
    // Shows how to assign Workbook.Settings.CultureInfo inside a custom LightCellsDataProvider, copy cell values and styles, and export the workbook with OoxmlSaveOptions in LightCells mode so numbers follow the chosen locale (e.g., German separators).
    public class CultureAwareLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Workbook _sourceWorkbook;
        private readonly CultureInfo _culture;
        private int _currentRow = -1;
        private int _currentColumn = -1;
        private int _maxRow;
        private int _maxColumn;

        public CultureAwareLightCellsDataProvider(Workbook sourceWorkbook, CultureInfo culture)
        {
            _sourceWorkbook = sourceWorkbook;
            _culture = culture;

            // Determine the used range of the first worksheet.
            Cells cells = _sourceWorkbook.Worksheets[0].Cells;
            _maxRow = cells.MaxDataRow;
            _maxColumn = cells.MaxDataColumn;
        }

        // No need to gather string data separately.
        public bool IsGatherString() => false;

        // Only one sheet is processed.
        public int SheetCount => 1;

        // Called before processing a sheet. Set the desired culture here.
        public bool StartSheet(int sheetIndex)
        {
            // Apply the culture to the workbook settings.
            _sourceWorkbook.Settings.CultureInfo = _culture;
            // Reset row/column counters.
            _currentRow = -1;
            _currentColumn = -1;
            // Process only the first sheet.
            return sheetIndex == 0;
        }

        // Move to the next row. Return -1 when no more rows.
        public int NextRow()
        {
            if (_currentRow < _maxRow)
            {
                _currentRow++;
                _currentColumn = -1;
                return _currentRow;
            }
            return -1;
        }

        // Called when a new row starts.
        public void StartRow(Row row) => _currentColumn = -1;

        // Move to the next cell in the current row. Return -1 when no more cells.
        public int NextCell()
        {
            if (_currentColumn < _maxColumn)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1;
        }

        // Populate the cell with the value from the source workbook.
        public void StartCell(Cell cell)
        {
            Cell srcCell = _sourceWorkbook.Worksheets[0].Cells[_currentRow, _currentColumn];
            // Preserve the original value and style.
            cell.PutValue(srcCell.Value);
            cell.SetStyle(srcCell.GetStyle());
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add sample numeric data.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cell cell = ws.Cells["A1"];
            cell.PutValue(12345.67); // Sample number.

            // Apply a number format that will be affected by culture settings.
            Style style = wb.CreateStyle();
            style.Custom = "#,##0.00"; // Uses group and decimal separators.
            cell.SetStyle(style);

            // Define the target culture (e.g., German uses comma as decimal separator).
            CultureInfo targetCulture = new CultureInfo("de-DE");

            // Create save options for XLSX and assign the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CultureAwareLightCellsDataProvider(wb, targetCulture)
            };

            // Save the workbook using the LightCells mode with culture‑specific formatting.
            string outputPath = "CultureSpecificNumberFormat.xlsx";
            wb.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with culture '{targetCulture.Name}'.");
        }
    }
}
