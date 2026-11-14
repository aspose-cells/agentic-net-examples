// Title: Delete blank columns on the first worksheet while updating references using DeleteOptions.UpdateReference in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete blank columns and keep formulas updated | How to use DeleteOptions.UpdateReference to remove empty columns from the first sheet | Aspose.Cells DeleteBlankColumns example with reference preservation | Remove all empty columns from a worksheet while updating cell references in .NET | Delete blank columns in an Excel file using Aspose.Cells and maintain formula links
// Tags: DeleteBlankColumns with UpdateReference | first worksheet column cleanup Aspose.Cells | preserve formulas after column deletion C# | Aspose.Cells DeleteOptions usage | blank column removal from Excel using .NET

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (index 0)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Configure delete options to update references after deletion
Aspose.Cells.DeleteOptions deleteOptions = new Aspose.Cells.DeleteOptions();
deleteOptions.UpdateReference = true;

// Delete all blank columns in the worksheet using the specified options
sheet.Cells.DeleteBlankColumns(deleteOptions);

// Save the modified workbook
workbook.Save("output.xlsx");
