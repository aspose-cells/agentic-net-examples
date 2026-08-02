using System;
using System.Collections;
using Aspose.Cells;

class KeywordCellCounter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Apple");
        cells["B1"].PutValue("Banana");
        cells["A2"].PutValue("Green Apple");
        cells["B2"].PutValue("Orange");
        cells["A3"].PutValue("Pineapple");
        cells["B3"].PutValue("Grape");

        // Keyword to search for
        string keyword = "Apple";

        // Counter for cells containing the keyword
        int matchCount = 0;

        // Enumerate all instantiated cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Get the displayed string value of the cell
            string cellText = cell.StringValue;

            // Check if the cell text contains the keyword (case‑insensitive)
            if (!string.IsNullOrEmpty(cellText) &&
                cellText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                matchCount++;
            }
        }

        Console.WriteLine($"Cells containing \"{keyword}\": {matchCount}");

        // Save the workbook
        workbook.Save("KeywordCount.xlsx");
    }
}