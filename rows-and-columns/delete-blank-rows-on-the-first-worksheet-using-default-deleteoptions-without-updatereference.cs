// Title: Delete blank rows from the first worksheet in an Excel workbook using Aspose.Cells for .NET with default DeleteOptions and no formula reference updates
// Common Searches: Aspose.Cells C# delete blank rows without updating formula references | How to use DeleteOptions to remove empty rows from the first sheet in .NET | DeleteBlankRows default settings keep cell references unchanged Aspose.Cells | Remove empty rows from worksheet while preserving formulas using Aspose.Cells | C# code to delete blank rows in Excel file without affecting formulas
// Tags: Aspose.Cells DeleteBlankRows with DeleteOptions | delete blank rows default DeleteOptions | disable formula reference update Aspose.Cells | first worksheet blank row removal C# | Excel workbook blank row deletion Aspose.Cells

// Load an existing workbook (replace with your file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (index 0)
var worksheet = workbook.Worksheets[0];

// Create DeleteOptions with default settings and disable reference updating
var deleteOptions = new Aspose.Cells.DeleteOptions();
deleteOptions.UpdateReference = false;   // Do not adjust formulas/references after deletion

// Delete all blank rows on the worksheet using the specified options
worksheet.Cells.DeleteBlankRows(deleteOptions);

// Save the modified workbook (replace with desired output path)
workbook.Save("output.xlsx");
