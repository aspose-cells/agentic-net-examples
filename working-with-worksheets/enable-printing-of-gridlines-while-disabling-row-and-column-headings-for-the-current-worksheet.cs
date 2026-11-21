// Title: How to print gridlines and hide row/column headings in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# print gridlines without row and column headings | Set PrintGridlines true and PrintHeadings false in Aspose.Cells | Disable printed headings while enabling gridlines in Excel with Aspose.Cells .NET | Configure worksheet page setup to show gridlines but hide headings using Aspose.Cells | C# Aspose.Cells example for printing only gridlines in a workbook
// Tags: Aspose.Cells PageSetup PrintGridlines | Aspose.Cells PageSetup PrintHeadings | C# Excel gridlines printing Aspose.Cells | hide row column headings Aspose.Cells | worksheet print settings Aspose.Cells .NET

// Create a new workbook (or load an existing one)
// Assuming the create rule provides a Workbook instance named "workbook"
var workbook = new Aspose.Cells.Workbook(); // create

// Access the first worksheet (or the current worksheet you want to modify)
var worksheet = workbook.Worksheets[0]; // current worksheet

// Enable printing of gridlines
worksheet.PageSetup.PrintGridlines = true;

// Disable printing of row and column headings
worksheet.PageSetup.PrintHeadings = false;

// Save the workbook (assuming the save rule uses the same "workbook" variable)
workbook.Save("Result.xlsx"); // save
