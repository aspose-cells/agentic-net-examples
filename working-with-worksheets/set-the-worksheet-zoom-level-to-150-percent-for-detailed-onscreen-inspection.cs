// Title: How to set a worksheet’s zoom level to 150 % using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a workbook with Aspose.Cells, sets the first worksheet’s Zoom property to 150, and saves the file. | Show a minimal Aspose.Cells example that changes the view scale of a worksheet to 150 % for detailed inspection. | Demonstrate how to programmatically adjust Excel sheet zoom to 150 % in C# using the Aspose.Cells Worksheet.Zoom property.
// Common Searches: Aspose.Cells C# set worksheet zoom to 150 percent example | How to change Excel sheet view scale with Aspose.Cells .NET | C# code to increase worksheet zoom level using Aspose.Cells API
// Tags: Aspose.Cells worksheet zoom property C# | set worksheet view scale Aspose.Cells | adjust Excel sheet zoom Aspose.Cells .NET | programmatic zoom level Excel workbook C# | Aspose.Cells zoom 150 percent example

using Aspose.Cells;

// Create a new workbook (or load an existing one)
// Here we create a new workbook; replace with load rule if needed
Workbook workbook = new Workbook();

// Access the first worksheet in the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Set the worksheet zoom level to 150%
worksheet.Zoom = 150;

// Save the workbook to a file
workbook.Save("ZoomedWorkbook.xlsx");
