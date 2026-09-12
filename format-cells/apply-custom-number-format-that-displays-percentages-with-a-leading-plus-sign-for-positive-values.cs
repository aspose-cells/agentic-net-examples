// Title: Format percentages with a leading plus sign for positive values using Aspose.Cells for .NET
// AI Prompts: Create a custom number format string in Aspose.Cells that displays '+' for positive percentages, '-' for negatives, and no sign for zero, then assign it to a style object. | Apply the custom percentage style to cells A1 and A2 in a new workbook and save the file as PercentWithPlus.xlsx.
// Common Searches: Aspose.Cells C# custom number format to show plus sign for positive percentages | How to display positive percentages with a leading '+' in an Excel file using Aspose.Cells | Set custom percentage style with positive/negative/zero formatting in .NET Aspose.Cells | Apply custom number format to specific cells and save workbook Aspose.Cells C# | Formatting percentages with sign symbols in Aspose.Cells example
// Tags: custom percentage number format Aspose.Cells | apply plus sign style to cells C# | percentage formatting with positive sign .xlsx | set custom style workbook Aspose.Cells | format cells as percentage with sign Aspose.Cells

using Aspose.Cells;

// CREATE_WORKBOOK
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Sample data: positive and negative percentages
sheet.Cells["A1"].PutValue(0.25);   // 25%
sheet.Cells["A2"].PutValue(-0.12);  // -12%

// Create a style with a custom number format that shows a leading plus sign for positives
Style percentStyle = workbook.CreateStyle();
percentStyle.Custom = "+0.00%;-0.00%;0.00%"; // positive;negative;zero

// Apply the style to the cells
sheet.Cells["A1"].SetStyle(percentStyle);
sheet.Cells["A2"].SetStyle(percentStyle);

// SAVE_WORKBOOK
workbook.Save("PercentWithPlus.xlsx");
