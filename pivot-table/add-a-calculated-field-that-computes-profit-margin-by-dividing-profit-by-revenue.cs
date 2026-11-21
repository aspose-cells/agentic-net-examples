// Title: Create a Profit Margin calculated column in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a profit margin column with a formula in Excel using Aspose.Cells C# | Aspose.Cells C# calculate profit divided by revenue for each row | set cell formula programmatically in Aspose.Cells to compute profit margin | generate derived column profit over revenue in an existing workbook with Aspose.Cells .NET | apply custom calculation to Excel sheet and save with Aspose.Cells C#
// Tags: Aspose.Cells profit margin formula | C# add calculated column Excel | Aspose.Cells set cell formula programmatically | Excel workbook compute profit revenue ratio .NET | auto compute derived field Aspose.Cells

// Load an existing workbook (replace "input.xlsx" with your file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (or specify by name)
var worksheet = workbook.Worksheets[0];

// Define the range where profit and revenue data exist.
// Example: profit in column B (index 1), revenue in column C (index 2), starting from row 2 (index 1)
int startRow = 1; // zero‑based index for row 2
int endRow = worksheet.Cells.MaxDataRow; // last row with data

// Add a calculated field "Profit Margin" in column D (index 3)
int profitMarginColumn = 3; // column D

// Set header for the calculated field
worksheet.Cells[0, profitMarginColumn].PutValue("Profit Margin");

// Fill the calculated field with the formula profit / revenue for each row
for (int row = startRow; row <= endRow; row++)
{
    // Formula: =B{row+1}/C{row+1}
    string formula = $"=B{row + 1}/C{row + 1}";
    worksheet.Cells[row, profitMarginColumn].Formula = formula;
}

// Optionally, calculate the formulas immediately
workbook.CalculateFormula();

// Save the updated workbook (replace "output.xlsx" with desired output path)
workbook.Save("output.xlsx");
