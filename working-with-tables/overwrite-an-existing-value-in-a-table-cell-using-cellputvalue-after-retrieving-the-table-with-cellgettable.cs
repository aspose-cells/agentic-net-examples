// Title: Overwrite a cell value in an Excel ListObject table using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells C# to locate a ListObject via Cell.GetTable and replace a specific data cell with a new string using the PutValue method. | Write C# code that retrieves the table containing cell B2, determines the first data row, and updates the value in the second column of that row. | Show how to modify an Excel workbook by changing a table cell value and then saving the file with Aspose.Cells.
// Common Searches: asp.net c# how to update a value in an Excel ListObject table using Aspose.Cells | asp.net example of modifying a table cell with Aspose.Cells and saving the workbook | replace data in first data row second column of an Excel table programmatically | c# code to get the table that contains a specific cell and change its content | asp.net modify Excel table cell and persist changes with Aspose.Cells
// Tags: Aspose.Cells GetTable API | Aspose.Cells PutValue method | modify ListObject cell C# | update Excel table cell Aspose | overwrite Excel table data .NET

// Load an existing workbook
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (adjust index or name as needed)
var worksheet = workbook.Worksheets[0];

// Get a cell that belongs to the target table (e.g., cell B2)
var cell = worksheet.Cells["B2"];

// Retrieve the table (ListObject) that contains this cell
var table = cell.GetTable();

if (table != null)
{
    // Determine the first data row and column of the table
    int dataStartRow = table.DataRange.FirstRow;
    int dataStartColumn = table.DataRange.FirstColumn;

    // Example: overwrite the value in the first data row, second column of the table
    // Adjust offsets as required for the specific cell you want to modify
    var targetCell = worksheet.Cells[dataStartRow, dataStartColumn + 1];

    // Overwrite the existing value
    targetCell.PutValue("New Value");
}

// Save the modified workbook
workbook.Save("output.xlsx");
