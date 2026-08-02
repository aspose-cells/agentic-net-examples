using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRegexDateSearch
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data in J1:J30 (some dates in dd/MM/yyyy format and other text)
            for (int i = 0; i < 30; i++)
            {
                // Alternate between a valid date string and other text
                string value = (i % 3 == 0) ? $"15/0{(i % 12) + 1}/2023" : $"Sample{i}";
                cells[0 + i, 9].PutValue(value); // Column J = index 9
            }

            // Prepare find options for regex search
            FindOptions options = new FindOptions
            {
                // Enable regex processing
                RegexKey = true,
                // Search only the cell values (not formulas)
                LookInType = LookInType.Values,
                // Exact match of the whole cell content (performance)
                LookAtType = LookAtType.EntireContent
            };

            // Restrict the search area to J1:J30
            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // J1
                StartColumn = 9,   // Column J
                EndRow = 29,       // J30
                EndColumn = 9
            };
            options.SetRange(searchArea);

            // Regular expression for dates in dd/MM/yyyy format
            string datePattern = @"\b\d{2}/\d{2}/\d{4}\b";

            // Loop to find all matching cells
            Cell previous = null;
            while (true)
            {
                Cell found = cells.Find(datePattern, previous, options);
                if (found == null)
                    break; // No more matches

                // Highlight the found cell (e.g., yellow background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = System.Drawing.Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                found.SetStyle(style);

                // Continue searching after the current cell
                previous = found;
            }

            // Save the workbook
            workbook.Save("RegexDateSearchResult.xlsx");
        }
    }
}