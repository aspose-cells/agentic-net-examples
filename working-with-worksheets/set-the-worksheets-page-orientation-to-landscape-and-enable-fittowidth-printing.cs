// Title: Set worksheet page orientation to landscape and enable fit‑to‑width printing with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that changes a worksheet's page orientation to landscape and configures FitToPagesWide = 1, FitToPagesTall = 0 using Aspose.Cells. | Show how to apply Aspose.Cells PageSetup properties to set landscape orientation and enable fit‑to‑width printing, then save the workbook. | Create a .NET example that creates a workbook, sets the first sheet to landscape, fits the content to one page wide, and writes the file.
// Common Searches: Aspose.Cells C# set worksheet page orientation to landscape | Fit to width printing Aspose.Cells page setup example | How to configure Aspose.Cells workbook to print on one page wide | Aspose.Cells set landscape orientation and fit to page width .NET | FitToPagesWide 1 FitToPagesTall 0 Aspose.Cells C#
// Tags: set worksheet landscape orientation Aspose.Cells | fit-to-width printing Aspose.Cells C# | Aspose.Cells page setup fit to pages wide | configure print options Aspose.Cells workbook | Aspose.Cells landscape page setup example

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (or add a new one if needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Set page orientation to Landscape
sheet.PageSetup.Orientation = Aspose.Cells.PageOrientationType.Landscape;

// Enable Fit‑to‑Width printing (fit to 1 page wide, unlimited height)
sheet.PageSetup.FitToPagesWide = 1;
sheet.PageSetup.FitToPagesTall = 0;   // 0 means "as many pages as needed" (fit to width only)

// Save the workbook
workbook.Save("Output.xlsx");
