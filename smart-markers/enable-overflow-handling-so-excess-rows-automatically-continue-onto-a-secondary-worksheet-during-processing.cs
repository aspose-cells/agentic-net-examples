// Title: C# Aspose.Cells example: automatically split a large DataTable across multiple worksheets when row limit is exceeded
// Common Searches: how to split a DataTable into multiple Excel worksheets using Aspose.Cells C# when rows exceed 1000 | Aspose.Cells C# create new worksheet automatically after reaching max rows per sheet | export large DataTable to Excel with automatic sheet overflow handling in .NET | C# write Excel file with header repeated on each new sheet using Aspose.Cells | Aspose.Cells example for dynamic worksheet creation based on row count limit
// Tags: Aspose.Cells automatic worksheet overflow | C# split DataTable across multiple Excel sheets | max rows per worksheet handling Aspose.Cells | header replication on new sheets Aspose.Cells | dynamic sheet naming Aspose.Cells C# | export large dataset to Excel with sheet split

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Define the maximum number of rows per worksheet (excluding the header row)
int maxRowsPerSheet = 1000;

// Sample data source (replace with actual data source)
System.Data.DataTable dataTable = GetDataTable(); // Assume this method returns the data to be written

// Keep track of the current worksheet and the row index within that sheet
Aspose.Cells.Worksheet currentSheet = workbook.Worksheets[0];
int currentRow = 0;

// Write header row once per worksheet
void WriteHeader(Aspose.Cells.Worksheet sheet)
{
    // Example: write column names from the DataTable
    for (int col = 0; col < dataTable.Columns.Count; col++)
    {
        sheet.Cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
    }
}

// Initialize the first sheet with a header
WriteHeader(currentSheet);
currentRow = 1; // Start writing data after the header

// Iterate through the data rows
foreach (System.Data.DataRow dr in dataTable.Rows)
{
    // If the current sheet has reached its row limit, create a new sheet and reset the row counter
    if (currentRow >= maxRowsPerSheet)
    {
        // Create a new worksheet with a unique name
        string newSheetName = $"Sheet{workbook.Worksheets.Count + 1}";
        currentSheet = workbook.Worksheets.Add(newSheetName);

        // Write header on the new sheet
        WriteHeader(currentSheet);
        currentRow = 1; // Reset to the first data row after the header
    }

    // Write the data row to the current worksheet
    for (int col = 0; col < dataTable.Columns.Count; col++)
    {
        currentSheet.Cells[currentRow, col].PutValue(dr[col]);
    }

    currentRow++;
}

// Save the workbook to a file
workbook.Save("ProcessedData.xlsx");

// ---------------------------------------------------------------------------
// Helper method to obtain the source DataTable (replace with actual implementation)
System.Data.DataTable GetDataTable()
{
    // Placeholder implementation: create a dummy table with sample data
    var dt = new System.Data.DataTable();
    dt.Columns.Add("ID", typeof(int));
    dt.Columns.Add("Name", typeof(string));
    dt.Columns.Add("Value", typeof(double));

    // Populate with sample rows (e.g., 2500 rows to trigger overflow)
    for (int i = 1; i <= 2500; i++)
    {
        dt.Rows.Add(i, $"Item {i}", i * 1.1);
    }

    return dt;
}
