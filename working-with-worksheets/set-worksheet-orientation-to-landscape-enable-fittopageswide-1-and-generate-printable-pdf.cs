// Title: Configure worksheet to landscape orientation and fit to one page width when exporting to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet orientation to landscape and fit to one page wide for PDF export | How to use PageSetup in Aspose.Cells to generate a PDF that fits a single page width | C# Aspose.Cells PDF export with landscape page setup and FitToPagesWide = 1 | Saving Excel worksheet as PDF with landscape orientation and automatic height scaling using Aspose.Cells | Fit printed worksheet to one page wide in Aspose.Cells .NET PDF conversion
// Tags: landscape orientation page setup Aspose.Cells | FitToPagesWide PDF export Aspose.Cells | worksheet scaling to single page width C# | Aspose.Cells PDF conversion page layout | C# set page orientation landscape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Create a new workbook (or load an existing one)
Workbook workbook = new Workbook(); // creates a new workbook with a default worksheet

// Access the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set page orientation to Landscape
sheet.PageSetup.Orientation = PageOrientationType.Landscape;

// Fit the printed sheet to 1 page wide (height will be scaled automatically)
sheet.PageSetup.FitToPagesWide = 1;
sheet.PageSetup.FitToPagesTall = 0; // 0 means automatic

// Save the workbook as a printable PDF
workbook.Save("Output.pdf", SaveFormat.Pdf);
