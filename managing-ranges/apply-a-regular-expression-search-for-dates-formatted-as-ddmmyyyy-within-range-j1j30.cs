using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class RegexDateSearch
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // create workbook
        Worksheet sheet = workbook.Worksheets[0];         // get first worksheet

        // Example data in J1:J30 (replace with your actual data or load a file)
        for (int i = 0; i < 30; i++)
        {
            // Populate some cells with date strings and other text for demonstration
            string value = (i % 5 == 0) ? $"15/0{(i % 12) + 1}/2023" : $"Sample{i}";
            sheet.Cells[i, 9].PutValue(value); // column J is index 9
        }

        // Configure find options for regex search
        FindOptions options = new FindOptions
        {
            LookInType = LookInType.Values,          // search in cell values
            LookAtType = LookAtType.EntireContent,   // exact match of the regex pattern
            RegexKey = true                          // treat the search key as a regular expression
        };

        // Regular expression for dates in dd/MM/yyyy format
        string datePattern = @"\b\d{2}/\d{2}/\d{4}\b";

        // Search within the specified range J1:J30
        Cell previousCell = null;
        Cell foundCell;

        // Loop to find all matching cells
        while ((foundCell = sheet.Cells.Find(datePattern, previousCell, options)) != null)
        {
            // Apply a highlight style to the found cell
            Style highlight = workbook.CreateStyle();
            highlight.ForegroundColor = Color.Yellow;
            highlight.Pattern = BackgroundType.Solid;
            foundCell.SetStyle(highlight);

            // Continue searching after the current cell
            previousCell = foundCell;
        }

        // Save the workbook with the highlighted results
        workbook.Save("RegexDateSearchResult.xlsx");
    }
}