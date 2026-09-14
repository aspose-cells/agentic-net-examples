// Title: Configure a worksheet for unlimited printed pages by setting FitToPagesWide and FitToPagesTall to zero with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets worksheet.PageSetup.FitToPagesWide = 0 and FitToPagesTall = 0 using Aspose.Cells and saves the workbook. | Provide a snippet that disables page scaling limits in Aspose.Cells by assigning zero to the FitToPagesWide and FitToPagesTall properties. | Write a method that creates a new workbook, configures unlimited printed page dimensions, and writes output.xlsx with Aspose.Cells.
// Common Searches: how to remove page fit constraints in Aspose.Cells C# | Aspose.Cells example for unlimited columns per printed page | C# code to set FitToPagesTall zero for full sheet printing | Aspose.Cells page setup to print entire worksheet without scaling | unlimited page width and height settings in Aspose.Cells
// Tags: Aspose.Cells page setup unlimited scaling | set FitToPagesWide to zero Aspose.Cells | set FitToPagesTall to zero Aspose.Cells | disable page fit constraints Aspose.Cells | print all rows on a single sheet Aspose.Cells

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Get the first worksheet (or any worksheet you need)
var worksheet = workbook.Worksheets[0];

// Set unlimited page width and height by assigning zero
worksheet.PageSetup.FitToPagesWide = 0;   // Unlimited columns per page
worksheet.PageSetup.FitToPagesTall = 0;   // Unlimited rows per page

// Save the workbook to a file
workbook.Save("output.xlsx");
