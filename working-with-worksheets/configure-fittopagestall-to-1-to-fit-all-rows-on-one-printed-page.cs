// Title: How to set FitToPagesTall = 1 in Aspose.Cells for .NET to print all rows on a single page
// Common Searches: Aspose.Cells C# set FitToPagesTall to 1 for single-page row printing | C# page setup fit all rows to one page using Aspose.Cells | Print worksheet with all rows on one page Aspose.Cells .NET example | Configure FitToPagesTall property in Aspose.Cells workbook | Aspose.Cells FitToPagesWide 0 and FitToPagesTall 1 usage
// Tags: FitToPagesTall property Aspose.Cells .NET | page setup fit rows single page | worksheet print scaling Aspose.Cells | FitToPagesWide zero automatic width | export workbook single page layout

// Create a new workbook (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Example data (optional, just to have something to print)
sheet.Cells["A1"].PutValue("Header");
sheet.Cells["A2"].PutValue("Row 1");
sheet.Cells["A3"].PutValue("Row 2");
sheet.Cells["A4"].PutValue("Row 3");

// Configure the page setup to fit all rows on a single printed page
sheet.PageSetup.FitToPagesTall = 1;   // Fit all rows (tall) to 1 page
sheet.PageSetup.FitToPagesWide = 0;   // Keep the original width (0 = automatic)

// Save the workbook to a file
workbook.Save("FitToPagesTallExample.xlsx");
