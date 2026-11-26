// Title: Set worksheet bottom margin to 2 points with Aspose.Cells for .NET to prevent text clipping
// Common Searches: Aspose.Cells C# set worksheet bottom margin to 2 points | How to avoid bottom text clipping in Excel using Aspose.Cells | Change page setup bottom margin in .NET Aspose.Cells example | Set Excel worksheet bottom margin in points with Aspose.Cells API | Adjust bottom margin to prevent clipping when exporting Excel via Aspose
// Tags: Aspose.Cells bottom margin configuration | worksheet page setup margin adjustment .NET | prevent Excel bottom clipping Aspose | set page margins programmatically C# | Excel bottom margin points Aspose.Cells

// Load an existing workbook (using the provided load rule)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (or any target worksheet)
var worksheet = workbook.Worksheets[0];

// Set the bottom margin to 2 points to avoid clipping (using the PageSetup property)
worksheet.PageSetup.BottomMargin = 2;

// Save the modified workbook (using the provided save rule)
workbook.Save("output.xlsx");
