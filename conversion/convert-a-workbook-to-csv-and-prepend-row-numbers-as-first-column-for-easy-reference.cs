// Title: Add Row Numbers to an Excel Sheet and Export as CSV using Aspose.Cells for .NET
// Description: A concise C# example that loads an existing workbook, inserts a new left‑most column, fills it with 1‑based row numbers up to the worksheet's MaxDataRow, and saves the result directly as a CSV file with Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | Excel to CSV conversion | add row numbers | insert first column | MaxDataRow | CSV export | worksheet manipulation | row index column
// Common Searches: Aspose.Cells add leading row number column before CSV export | C# insert column at position 0 in Excel worksheet | How to write row numbers to Excel using Aspose.Cells | Export Excel to CSV with row index column .NET | Determine last data row in Aspose.Cells and number rows
// Developer Intent: Insert a sequential row‑number column at the start of a worksheet and generate a CSV file that includes this identifier.
// Use Cases: Create CSV reports that retain a reference to the original Excel row numbers. | Provide a unique identifier for each line when importing data into systems that require a primary key. | Generate debug‑friendly CSV dumps where row ordering is explicitly shown.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a first column, populates it with 1‑based row numbers, and saves the sheet as CSV. | Explain how Cells.MaxDataRow can be used to limit row‑number insertion to rows that contain data. | Suggest a way to skip completely empty rows while adding the row‑number column before exporting to CSV.

using System;
using Aspose.Cells;

// A concise C# example that loads an existing workbook, inserts a new left‑most column, fills it with 1‑based row numbers up to the worksheet's MaxDataRow, and saves the result directly as a CSV file with Aspose.Cells.
class WorkbookToCsvWithRowNumbers
{
    static void Main()
    {
        // Paths for source workbook and destination CSV
        string sourcePath = "input.xlsx";
        string destinationPath = "output.csv";

        // Load the existing workbook (create + load lifecycle)
        Workbook workbook = new Workbook(sourcePath);

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at the very left (index 0)
        cells.InsertColumn(0);

        // Determine the last row that contains data
        int lastDataRow = cells.MaxDataRow;

        // Populate the new first column with row numbers (1‑based)
        for (int row = 0; row <= lastDataRow; row++)
        {
            cells[row, 0].PutValue(row + 1);
        }

        // Save the modified workbook as CSV (save lifecycle)
        workbook.Save(destinationPath, SaveFormat.Csv);
    }
}
