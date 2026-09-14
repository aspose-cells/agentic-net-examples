// Title: Enable Page Break Preview mode for a worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets a worksheet’s ViewType to PageBreakPreview using Aspose.Cells and saves the workbook. | Show how to switch an existing worksheet to page break preview view before exporting to Excel with Aspose.Cells. | Provide a snippet that toggles the view mode of a specific worksheet to PageBreakPreview and writes the file to disk.
// Common Searches: Aspose.Cells C# set worksheet view to page break preview programmatically | How to display page break locations in an Excel file using Aspose.Cells .NET | C# change worksheet view type to PageBreakPreview before saving workbook | Enable page break preview mode for a specific sheet with Aspose.Cells API | Aspose.Cells ViewType PageBreakPreview example code
// Tags: Aspose.Cells set worksheet view type | C# page break preview mode | Aspose.Cells ViewType PageBreakPreview | save workbook after view change | Excel page break visualization Aspose.Cells

// Create a new workbook (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(); // create

// Access the first worksheet (or any specific worksheet)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Enable Page Break Preview mode for the worksheet
sheet.ViewType = Aspose.Cells.ViewType.PageBreakPreview;

// Save the workbook to a file
workbook.Save("PageBreakPreview.xlsx"); // save
