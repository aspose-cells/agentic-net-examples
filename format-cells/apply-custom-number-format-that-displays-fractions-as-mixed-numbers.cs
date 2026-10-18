// Title: Applying a custom mixed‑fraction number format (# ??/??) to an Excel cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format to display mixed fractions | How to format a cell as 2 3/4 using # ??/?? pattern in Aspose.Cells .NET | C# example for showing fractions as mixed numbers in Excel with Aspose.Cells | Apply mixed fraction number format to a worksheet cell using Aspose.Cells API | Custom number format for mixed fractions in an Aspose.Cells workbook C#
// Tags: Aspose.Cells custom number format mixed fraction | C# set cell style number pattern | Excel mixed fraction format Aspose.Cells | cell style custom format # ??/?? Aspose.Cells | save workbook with custom format xlsx Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Put a numeric value that will be displayed as a mixed fraction
Aspose.Cells.Cell cell = worksheet.Cells["A1"];
cell.PutValue(2.75); // Example value: 2 3/4

// Create a style and set a custom number format for mixed fractions
Aspose.Cells.Style style = cell.GetStyle();
style.Custom = "# ??/??"; // Displays numbers as mixed fractions (e.g., 2 3/4)

// Apply the style to the cell
cell.SetStyle(style);

// Save the workbook to a file
workbook.Save("MixedFractionExample.xlsx");
