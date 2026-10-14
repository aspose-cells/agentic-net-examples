// Title: Set custom page margins (0.5" top/bottom, 0.3" left/right) for an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook with Aspose.Cells, accesses a worksheet, and sets the PageSetup.TopMargin, BottomMargin, LeftMargin, and RightMargin to specific inch values before saving. | Show how to adjust print margins in inches for an Excel sheet via the Aspose.Cells PageSetup object in a .NET application. | Demonstrate configuring custom worksheet margins (0.5 inches top/bottom, 0.3 inches left/right) and exporting the file as an .xlsx using Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet top margin to 0.5 inches | How to change left and right page margins in Excel using Aspose.Cells .NET | Programmatically define custom print margins for Excel file with Aspose.Cells | C# Aspose.Cells PageSetup margin values in inches example | Set custom page layout margins for Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells PageSetup margin configuration | C# set Excel worksheet print margins | custom page margins inches Aspose.Cells | adjust worksheet page layout .NET | Excel workbook margin settings Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (or any specific worksheet)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Get the PageSetup object to modify page settings
Aspose.Cells.PageSetup pageSetup = worksheet.PageSetup;

// Set custom page margins (values are in inches)
pageSetup.TopMargin = 0.5;      // Top margin = 0.5 inches
pageSetup.BottomMargin = 0.5;   // Bottom margin = 0.5 inches
pageSetup.LeftMargin = 0.3;     // Left margin = 0.3 inches
pageSetup.RightMargin = 0.3;    // Right margin = 0.3 inches

// Save the workbook to a file
workbook.Save("CustomMargins.xlsx");
