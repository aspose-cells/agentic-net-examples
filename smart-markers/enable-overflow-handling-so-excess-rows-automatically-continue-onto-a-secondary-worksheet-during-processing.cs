// Title: How to automatically overflow rows to additional worksheets in Aspose.Cells for .NET when exceeding Excel's 1,048,576 row limit
// AI Prompts: Write C# code using Aspose.Cells that iterates over an IEnumerable and creates a new worksheet each time the row index reaches 1,048,576, preserving previously written data. | Adjust the existing loop to reset the row counter and add a worksheet with an incremented name when the maximum rows per sheet are exceeded. | Create a reusable method SaveToExcelWithOverflow(IEnumerable<string> data, string filePath) that handles row overflow by automatically adding secondary worksheets.
// Common Searches: Aspose.Cells .NET how to split export into multiple sheets after reaching Excel row limit | C# write large collection to Excel and automatically create new worksheets when rows exceed 1,048,576 | handling Excel row overflow with Aspose.Cells without manual sheet management | auto paginate data across worksheets using Aspose.Cells for large datasets
// Tags: Aspose.Cells automatic worksheet overflow | Excel row limit handling with Aspose.Cells | dynamic sheet creation based on row count C# | large dataset export to multiple worksheets Aspose | overflow rows to new sheet Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet (primary sheet)
Aspose.Cells.Worksheet primarySheet = workbook.Worksheets[0];
primarySheet.Name = "DataSheet1";

// Excel 2007+ maximum rows per worksheet
const int MaxRowsPerSheet = 1048576;

// Sample data source (replace with actual data source)
System.Collections.Generic.List<string> data = GetSampleData(); // Assume this method returns a list of strings

int currentRow = 0;               // Row index within the current sheet
int sheetIndex = 1;               // Counter for secondary sheets
Aspose.Cells.Worksheet currentSheet = primarySheet;

// Iterate through the data and write each item to a new row
foreach (string item in data)
{
    // If the current row exceeds the maximum allowed rows, create a new worksheet
    if (currentRow >= MaxRowsPerSheet)
    {
        // Add a new worksheet and set it as the current sheet
        currentSheet = workbook.Worksheets.Add($"DataSheet{sheetIndex + 1}");
        sheetIndex++;
        currentRow = 0; // Reset row counter for the new sheet
    }

    // Write the data to column A of the current row
    currentSheet.Cells[currentRow, 0].PutValue(item);

    // Move to the next row
    currentRow++;
}

// Save the workbook (using the provided save rule)
workbook.Save("OverflowHandledWorkbook.xlsx");

// ---------------------------------------------------------------------------
// Helper method to simulate data retrieval (replace with real implementation)
System.Collections.Generic.List<string> GetSampleData()
{
    var list = new System.Collections.Generic.List<string>();
    // Generate more rows than a single sheet can hold to demonstrate overflow
    for (int i = 0; i < MaxRowsPerSheet + 5000; i++)
    {
        list.Add($"Row {i + 1}");
    }
    return list;
}
