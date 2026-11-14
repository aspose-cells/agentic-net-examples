// Title: Use Aspose.Cells for .NET to add an IF formula that shows a discount label only when the percentage is greater than zero
// Common Searches: asp.net aspose.cells c# add IF formula to display discount only when percentage > 0 | how to concatenate static text with a numeric cell value in an Aspose.Cells formula | programmatically set and calculate an Excel IF function using Aspose.Cells for .NET | generate Excel report with conditional discount label using Aspose.Cells C# API | evaluate formulas after setting them in an Aspose.Cells workbook before saving
// Tags: Aspose.Cells set IF formula programmatically | Excel conditional text output using Aspose.Cells | C# combine static text and cell value in Excel formula | run formula calculation in Aspose.Cells workbook | save workbook with conditional content Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Example: set a discount percentage in cell A1 (as a numeric value, e.g., 15%)
sheet.Cells["A1"].PutValue(15); // 15%

// In cell B1, use the IF function to display discount information only when the percentage is greater than zero
// The formula concatenates the text "Discount: " with the percentage value and a percent sign
sheet.Cells["B1"].Formula = @"IF(A1>0, ""Discount: "" & A1 & ""%"", """")";

// Calculate the formulas so that the result appears in B1
workbook.CalculateFormula();

// Save the workbook to a file (adjust the path as needed)
workbook.Save("DiscountReport.xlsx");
