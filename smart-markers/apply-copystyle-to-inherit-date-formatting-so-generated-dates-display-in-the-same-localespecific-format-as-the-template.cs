// Title: How to copy a locale‑specific date style from a template workbook and apply it to generated cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy date style from template workbook | Apply locale‑specific date format to new cells with Aspose.Cells | Set cell style based on another cell in Aspose.Cells .NET | Inherit date formatting from a template when generating Excel rows C# | Use Aspose.Cells to preserve template date format while populating dates
// Tags: copy cell style Aspose.Cells | apply locale date format Aspose.Cells | set style from source cell C# | inherit date formatting template Excel | populate dates with template style Aspose.Cells

// Load the template workbook (contains the desired locale‑specific date format)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("Template.xlsx");

// Access the first worksheet (adjust index/name as needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Assume cell A1 in the template holds the correct date style
Aspose.Cells.Cell styleSourceCell = sheet.Cells["A1"];
Aspose.Cells.Style dateStyle = styleSourceCell.GetStyle();

// Generate some dates starting from row 2
for (int row = 1; row <= 10; row++) // rows are zero‑based; row 1 = A2
{
    // Create a new cell in column A for the generated date
    Aspose.Cells.Cell dateCell = sheet.Cells[row, 0]; // column 0 = A

    // Set the cell value to a DateTime (example: today + offset)
    dateCell.PutValue(DateTime.Today.AddDays(row - 1));

    // Apply the copied style so the date appears in the same locale‑specific format
    dateCell.SetStyle(dateStyle);
}

// Save the workbook with the populated dates
workbook.Save("Output.xlsx");
