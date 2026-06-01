using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook culture to French (France) – uses comma as decimal separator
        wb.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Create a style with a custom numeric format; separators will follow the set culture
        Style style = wb.CreateStyle();
        style.Custom = "#,##0.00"; // French will display as "12 345,67"

        // Apply the style to a cell and put a numeric value
        Cell cell = wb.Worksheets[0].Cells["A1"];
        cell.PutValue(12345.67);
        cell.SetStyle(style);

        // Create a LightCellsDataProvider that writes additional data rows
        LightCellsDataProvider provider = new SimpleProvider();

        // Configure OoxmlSaveOptions to use the LightCellsDataProvider
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = provider
        };

        // Save the workbook with culture‑specific formatting
        wb.Save("CultureSpecific.xlsx", saveOptions);
    }

    // Minimal LightCellsDataProvider implementation
    class SimpleProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentColumn = -1;

        // Sample data to be written (column index, value)
        private readonly string[,] data = new string[,]
        {
            { "B1", "3000,50" }, // French formatted number as string
            { "B2", "4500,75" }
        };

        public bool IsGatherString() => true;

        public int SheetCount => 1;

        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        public int NextRow()
        {
            if (currentRow < data.GetLength(0) - 1)
            {
                currentRow++;
                currentColumn = -1;
                return currentRow;
            }
            return -1; // No more rows
        }

        public void StartRow(Row row) => currentColumn = -1;

        public int NextCell()
        {
            if (currentColumn < data.GetLength(1) - 1)
            {
                currentColumn++;
                return currentColumn;
            }
            return -1; // No more cells
        }

        public void StartCell(Cell cell)
        {
            // Write the string value; the workbook's culture will be used when the file is opened
            cell.PutValue(data[currentRow, currentColumn]);
        }
    }
}