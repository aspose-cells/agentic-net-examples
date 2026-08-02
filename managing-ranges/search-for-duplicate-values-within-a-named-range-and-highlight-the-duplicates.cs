// Title: Highlight Duplicate Values in a Named Range with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills column A with sample data, defines a named range called FruitList (A1:A6), and applies a DuplicateValues conditional‑formatting rule that colors repeated entries yellow. The file is saved as HighlightedDuplicates.xlsx.
// Keywords: Aspose.Cells duplicate detection C# | conditional formatting duplicate values .NET | create named range Aspose.Cells | Excel duplicate highlight Aspose | C# workbook conditional formatting example
// Common Searches: Aspose.Cells highlight duplicates in a range | C# conditional formatting for duplicate cells using Aspose | how to create and use named ranges with Aspose.Cells | duplicate value detection in Excel via Aspose.Cells .NET | apply yellow fill to repeated entries with Aspose.Cells
// Developer Intent: Add a DuplicateValues conditional‑formatting rule to a named range so repeated cells are highlighted.
// Use Cases: Detect and flag repeated product IDs during inventory uploads. | Identify duplicate email addresses before a mass‑mail campaign. | Highlight recurring error codes in a system‑log worksheet for quick troubleshooting. | Mark duplicated student IDs in a class roster to prevent enrollment mistakes.
// AI Prompts: Write C# code with Aspose.Cells that uses a red background instead of yellow for duplicate values in a named range. | Explain how to change the font color and add a border to cells flagged as duplicates. | Provide step‑by‑step instructions for creating a named range and applying DuplicateValues conditional formatting in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, fills column A with sample data, defines a named range called FruitList (A1:A6), and applies a DuplicateValues conditional‑formatting rule that colors repeated entries yellow. The file is saved as HighlightedDuplicates.xlsx.
class HighlightDuplicates
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate sample data with duplicates in column A.
            string[] data = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape" };
            for (int i = 0; i < data.Length; i++)
            {
                cells[i, 0].PutValue(data[i]); // A1, A2, ...
            }

            // Create a named range that covers the data (A1:A6).
            AsposeRange namedRange = cells.CreateRange("A1", "A6");
            namedRange.Name = "FruitList";

            // Add a conditional formatting rule to highlight duplicate values.
            int cfIndex = ws.ConditionalFormattings.Add();
            var cf = ws.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting using the named range bounds.
            CellArea area = new CellArea
            {
                StartRow = namedRange.FirstRow,
                EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                StartColumn = namedRange.FirstColumn,
                EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
            };
            cf.AddArea(area);

            // Add a condition of type DuplicateValues.
            int condIndex = cf.AddCondition(FormatConditionType.DuplicateValues);
            var condition = cf[condIndex];

            // Set the highlight style (yellow background) for duplicate cells.
            condition.Style.BackgroundColor = Color.Yellow;

            // Save the workbook.
            wb.Save("HighlightedDuplicates.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
