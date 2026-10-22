// Title: Resetting the worksheet print area after adding rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a given number of rows into a worksheet, removes the current print area, and defines a new print area that spans the entire used range using Aspose.Cells. | Demonstrate how to compute the updated used range after adding rows and set it as the worksheet's print area with the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# update print area after adding rows | Programmatically set Excel print area to used range with Aspose.Cells | How to recalculate print area after inserting rows in .NET Excel file | Remove old print area and apply new one using Aspose.Cells API
// Tags: Aspose.Cells reset print area C# | insert rows adjust print area .NET | used range calculation Aspose.Cells | clear and set worksheet print area programmatically | Excel print area after data changes Aspose.Cells

using Aspose.Cells;

// The example loads an existing workbook, inserts five rows after the last data row, optionally writes sample values, clears any existing print area, computes the new used range, sets the print area to that range, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (modify as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last row with data
        int lastDataRow = sheet.Cells.MaxDataRow;

        // Insert new rows after the existing data (example: 5 rows)
        sheet.Cells.InsertRows(lastDataRow + 1, 5);

        // Optionally populate the new rows with sample data
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[lastDataRow + 1 + i, 0].PutValue($"New Data {i + 1}");
        }

        // ----- Reset the print area -----
        // Clear any existing print area
        sheet.PageSetup.PrintArea = "";

        // Calculate the new used range
        int firstRow = sheet.Cells.MinRow;
        int firstCol = sheet.Cells.MinColumn;
        int lastRow = sheet.Cells.MaxDataRow;
        int lastCol = sheet.Cells.MaxDataColumn;

        // Build the address string for the new print area (e.g., A1:D20)
        string newPrintArea = CellArea.CreateCellArea(firstRow, firstCol, lastRow, lastCol).ToString();

        // Apply the new print area
        sheet.PageSetup.PrintArea = newPrintArea;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
