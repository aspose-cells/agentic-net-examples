using System;
using Aspose.Cells;

namespace HyperlinkWithLightCells
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up OoxmlSaveOptions to use a custom LightCellsDataProvider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomProvider()
            };

            // Save the workbook; the provider will stream cell data including hyperlinks
            workbook.Save("HyperlinkLightCellsDemo.xlsx", saveOptions);
        }
    }

    // Custom LightCellsDataProvider that writes data and embeds hyperlinks via formulas
    public class CustomProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentColumn = -1;
        private const int TotalRows = 6;    // 1 header + 5 data rows
        private const int TotalColumns = 3; // ID, Name, Price

        // Only process the first worksheet (index 0)
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index to be saved, or -1 when done
        public int NextRow()
        {
            currentRow++;
            currentColumn = -1;
            return currentRow < TotalRows ? currentRow : -1;
        }

        // No special row initialization needed
        public void StartRow(Row row) { }

        // Return the next column index for the current row, or -1 when the row is finished
        public int NextCell()
        {
            currentColumn++;
            return currentColumn < TotalColumns ? currentColumn : -1;
        }

        // Populate each cell; for the "Name" column we add a hyperlink using the HYPERLINK formula
        public void StartCell(Cell cell)
        {
            if (currentRow == 0) // Header row
            {
                switch (currentColumn)
                {
                    case 0: cell.PutValue("ID"); break;
                    case 1: cell.PutValue("Name"); break;
                    case 2: cell.PutValue("Price"); break;
                }
            }
            else // Data rows
            {
                int id = currentRow; // Simple ID based on row number
                switch (currentColumn)
                {
                    case 0:
                        cell.PutValue(id);
                        break;
                    case 1:
                        // Embed a hyperlink: display text "Item{id}", URL "https://example.com/item{id}"
                        string url = $"https://example.com/item{id}";
                        string display = $"Item{id}";
                        cell.Formula = $"HYPERLINK(\"{url}\",\"{display}\")";
                        break;
                    case 2:
                        // Example price value
                        double price = 10.0 + id;
                        cell.PutValue(price);
                        break;
                }
            }
        }

        // Indicate whether string values should be gathered into a global string pool
        public bool IsGatherString() => true;
    }
}