// Title: Set a custom column width for a specific table column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to set column B width to 40 characters in an Excel worksheet. | Show how to adjust a worksheet column width in pixels using Aspose.Cells for .NET. | Provide an example that writes a long text to a cell and then expands the column to fit the content with Aspose.Cells.
// Common Searches: Aspose.Cells C# set column width for column B to fit long text | How to specify column width in characters with Aspose.Cells .NET | Set Excel column width in pixels using Aspose.Cells C# example | Increase column width for a table column in Aspose.Cells workbook | Adjust column size to accommodate long strings in Aspose.Cells
// Tags: Aspose.Cells set column width C# | column width characters Aspose.Cells | column width pixels Aspose.Cells | custom column width Excel Aspose.Cells | adjust worksheet column size for long text

// Create a new workbook (or load an existing one)
// If you need to load an existing file, replace the constructor argument with the file path.
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (you can also add a new one if needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Example: put a long text string into column B (index 1) of row 1 (index 0)
sheet.Cells[0, 1].PutValue("This is a very long text string that needs extra column width to be fully visible.");

// Set a custom column width for column B (index 1)
// Width is measured in characters of the standard font. Adjust the value as needed.
sheet.Cells.SetColumnWidth(1, 40); // 40 characters wide

// Optionally, you can also set the column width in pixels:
// sheet.Cells.SetColumnWidthPixel(1, 300);

// Save the workbook to a file
workbook.Save("CustomColumnWidth.xlsx");
