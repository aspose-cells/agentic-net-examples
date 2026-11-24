// Title: Auto‑fit a range of columns and then set precise pixel widths using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit columns then set column width in pixels | Set exact pixel width for multiple columns after AutoFitColumn in Aspose.Cells | How to apply SetColumnWidthPixel to a column range after auto‑fitting in .NET | Aspose.Cells .NET example for adjusting column widths to specific pixel values | AutoFitColumn followed by SetColumnWidthPixel for precise column layout in C#
// Tags: auto-fit columns Aspose.Cells C# | set column width pixel Aspose.Cells | column pixel alignment .NET | adjust column widths after autofit C# | Aspose.Cells column width customization

// Load an existing workbook (replace with your actual file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("Input.xlsx");

// Get the first worksheet (or any worksheet you need)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Define the range of columns you want to auto‑fit and then set precise pixel widths
int startColumn = 0;   // Column A (zero‑based index)
int endColumn   = 4;   // Column E (exclusive)

// Auto‑fit the columns to their contents
for (int col = startColumn; col < endColumn; col++)
{
    sheet.AutoFitColumn(col);
}

// After auto‑fitting, set exact pixel widths for each column
// Adjust the pixel values as required for your layout
int[] pixelWidths = new int[] { 80, 120, 100, 150 }; // Widths for columns A‑D

for (int i = 0; i < pixelWidths.Length; i++)
{
    int colIndex = startColumn + i;
    sheet.Cells.SetColumnWidthPixel(colIndex, pixelWidths[i]);
}

// Save the modified workbook (replace with your desired output path)
workbook.Save("Output.xlsx");
