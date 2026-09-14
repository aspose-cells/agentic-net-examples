// Title: How to format a cell in Aspose.Cells for .NET to show scientific notation with three significant digits
// AI Prompts: Write C# code that assigns the custom number format "0.00E+00" to a cell in an Aspose.Cells workbook to display values in scientific notation with three significant digits. | Demonstrate setting a style on a specific worksheet cell using Aspose.Cells for .NET to render numbers in scientific notation with a precision of three significant digits.
// Common Searches: Aspose.Cells C# apply custom number format 0.00E+00 to a cell | How to display numbers in scientific notation with three significant figures using Aspose.Cells for .NET | C# example for formatting Excel cell as scientific notation with specific precision in Aspose.Cells | Set cell style to scientific notation with three significant digits in Aspose.Cells workbook
// Tags: custom number format scientific notation Aspose.Cells | apply 0.00E+00 format C# | cell style scientific notation .NET | three significant digits Excel formatting Aspose | Aspose.Cells workbook cell formatting example

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Put a numeric value into a cell
Aspose.Cells.Cell cell = sheet.Cells["A1"];
cell.PutValue(12345.6789);

// Apply a custom number format for scientific notation with three significant digits
Aspose.Cells.Style style = cell.GetStyle();
style.Custom = "0.00E+00"; // one digit before the decimal and two after = three significant digits
cell.SetStyle(style);

// Save the workbook to a file
workbook.Save("output.xlsx");
