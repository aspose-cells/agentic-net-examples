// Title: Apply an AutoFilter on an Aspose.Cells table to display rows with scores greater than a specified value (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject covering a range, fills it with ID and Score columns, and sets an AutoFilter on the Score column using a variable threshold before saving the file. | Write a C# method using Aspose.Cells that accepts a numeric threshold, applies a numeric >= filter to a ListObject's column, and outputs the filtered workbook.
// Common Searches: Aspose.Cells C# filter ListObject rows where numeric column is greater than 50 | C# Aspose.Cells auto filter table based on variable threshold | how to use Aspose.Cells AutoFilter with ListObject to show high scores
// Tags: Aspose.Cells ListObject AutoFilter C# | filter Excel table rows by numeric threshold Aspose.Cells | C# apply numeric criteria to Aspose.Cells ListObject | create Excel table with auto filter using Aspose.Cells | dynamic score threshold filter Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// Creates a workbook, adds a ListObject covering A1:B11, populates ID and Score data, applies an AutoFilter on the Score column to keep rows where Score >= 50, and saves the result as FilteredTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Score");

            // Data rows (ID, Score)
            for (int i = 2; i <= 11; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID
                sheet.Cells[i - 1, 1].PutValue((i - 1) * 10);       // Score
            }

            // Add a ListObject (table) covering A1:B11
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 11, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "DataTable";

            // Apply an AutoFilter to show rows where Score >= threshold
            double threshold = 50.0;
            // Column index 1 corresponds to the "Score" column (0‑based)
            // Use a criteria string for numeric comparison
            table.AutoFilter.Filter(1, $">={threshold}");

            // Save the workbook
            workbook.Save("FilteredTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
