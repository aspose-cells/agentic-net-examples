// Title: How to AutoFit newly inserted rows in an Excel worksheet using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# AutoFitRows after inserting rows in existing workbook | C# code to auto adjust row height for a range of rows after InsertRows | Maintain uniform row height when adding rows with Aspose.Cells | AutoFitRows method usage for newly added rows in Aspose.Cells C# example | Fit row height automatically after inserting rows with Aspose.Cells library
// Tags: AutoFitRows on inserted row range Aspose.Cells | InsertRows followed by AutoFitRows C# | Uniform row height Excel Aspose.Cells | Worksheet AutoFitRows method usage | Adjust row height after row insertion Aspose.Cells

using Aspose.Cells;

// Loads an existing Excel workbook, inserts three rows at index 5, populates them with sample data, applies Worksheet.AutoFitRows to the inserted range to ensure uniform row height, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Row index where new rows will be inserted (0‑based)
        int insertRowIndex = 5; // example: after row 5

        // Number of rows to insert
        int rowsToInsert = 3;

        // Insert the rows
        sheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

        // Fill the newly inserted rows with sample data
        for (int i = 0; i < rowsToInsert; i++)
        {
            int row = insertRowIndex + i;
            sheet.Cells[row, 0].PutValue($"Item {row + 1}");
            sheet.Cells[row, 1].PutValue((row + 1) * 10);
        }

        // AutoFit the rows that were added to ensure uniform height
        sheet.AutoFitRows(insertRowIndex, insertRowIndex + rowsToInsert - 1);

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx");
    }
}
