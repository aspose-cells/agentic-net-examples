// Title: Add Clickable Hyperlinks While Streaming Cells with LightCellsDataProvider in Aspose.Cells for .NET
// Description: Learn how to insert hyperlinks into specific cells before saving a workbook that is exported with a custom **LightCellsDataProvider**. The example creates a workbook, adds hyperlinks to the URL column, streams rows using `LightCellsDataProvider`, and saves the file with `OoxmlSaveOptions`. The hyperlinks are persisted in the resulting XLSX, providing low‑memory, high‑performance export for large datasets. See the full source on the [Aspose.Cells GitHub repository](https://github.com/aspose-cells/Aspose.Cells-for-.NET).
// Keywords: Aspose.Cells | LightCellsDataProvider | C# hyperlink | .NET Excel export | OoxmlSaveOptions | low memory Excel generation | streaming workbook | clickable links in XLSX | GitHub Aspose.Cells example
// Common Searches: add hyperlink with LightCellsDataProvider Aspose.Cells | C# stream Excel rows and keep hyperlinks | Aspose.Cells low memory export clickable URLs | how to use OoxmlSaveOptions with hyperlinks | Aspose.Cells LightCells example GitHub
// Developer Intent: Insert clickable hyperlinks into cells while streaming data with a custom LightCellsDataProvider and save the workbook efficiently.
// Use Cases: Export a massive product catalog where each website URL becomes a clickable link without loading the entire sheet into memory. | Generate a financial report with reference links to external documents, using LightCells for low‑memory processing. | Create an automated data pipeline that streams rows from a database and embeds navigation links for quick access in the final XLSX.
// AI Prompts: Show me how to modify CustomLightCellsDataProvider so that any cell containing a URL string automatically receives a hyperlink. | Provide C# code that adds hyperlinks to a range of cells after LightCellsDataProvider has written the data, ensuring they are saved correctly. | Explain how to programmatically verify that hyperlinks exist in the XLSX file when using LightCellsDataProvider and OoxmlSaveOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkWithLightCells
{
    // Custom LightCellsDataProvider that streams cell values to the workbook.
    // Learn how to insert hyperlinks into specific cells before saving a workbook that is exported with a custom **LightCellsDataProvider**. The example creates a workbook, adds hyperlinks to the URL column, streams rows using `LightCellsDataProvider`, and saves the file with `OoxmlSaveOptions`. The hyperlinks are persisted in the resulting XLSX, providing low‑memory, high‑performance export for large datasets. See the full source on the [Aspose.Cells GitHub repository](https://github.com/aspose-cells/Aspose.Cells-for-.NET).
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        // Sample data to be written.
        private readonly string[,] _data = new string[,]
        {
            { "ID", "Name", "Website" },
            { "1", "Aspose", "https://www.aspose.com" },
            { "2", "Google", "https://www.google.com" },
            { "3", "Microsoft", "https://www.microsoft.com" }
        };

        private int _currentRow = -1;
        private int _currentColumn = -1;

        // No additional sheets are needed; only the first sheet is processed.
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 when finished.
        public int NextRow()
        {
            _currentRow++;
            _currentColumn = -1;
            return _currentRow < _data.GetLength(0) ? _currentRow : -1;
        }

        // Called before writing cells of the current row.
        public void StartRow(Row row)
        {
            // No special row handling required.
        }

        // Return the next column index to be saved, or -1 when the row is finished.
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn < _data.GetLength(1) ? _currentColumn : -1;
        }

        // Called for each cell; set its value here.
        public void StartCell(Cell cell)
        {
            cell.PutValue(_data[_currentRow, _currentColumn]);
        }

        // Indicates whether string values should be gathered into a global string pool.
        public bool IsGatherString()
        {
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add hyperlinks to the cells that will contain website URLs.
            //    The hyperlinks are added before saving; they will be persisted together
            //    with the streamed cell data.
            //    Here we add a hyperlink to cell C2, C3 and C4 respectively.
            sheet.Hyperlinks.Add("C2", 1, 1, "https://www.aspose.com");
            sheet.Hyperlinks.Add("C3", 1, 1, "https://www.google.com");
            sheet.Hyperlinks.Add("C4", 1, 1, "https://www.microsoft.com");

            // 3. Create OoxmlSaveOptions and assign the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomLightCellsDataProvider()
            };

            // 4. Save the workbook using the light‑weight mode.
            workbook.Save("HyperlinksWithLightCells.xlsx", saveOptions);
        }
    }
}
