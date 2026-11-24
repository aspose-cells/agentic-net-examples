// Title: Configure unlimited page scaling in Aspose.Cells for .NET by setting FitToPagesWide and FitToPagesTall to 0
// Common Searches: Aspose.Cells C# set FitToPagesWide to 0 for no page width limit | How to remove page height restriction with FitToPagesTall in Aspose.Cells .NET | Unlimited printing area using PageSetup in Aspose.Cells workbook | C# Aspose.Cells page setup auto scaling without fit-to-page constraints | Set worksheet page scaling to automatic in Aspose.Cells for Excel export
// Tags: Aspose.Cells page setup unlimited scaling | FitToPagesWide zero configuration | FitToPagesTall zero configuration | C# worksheet printing without page limits | Aspose.Cells auto page scaling .NET | Excel export page setup Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (or any worksheet you need)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Get the PageSetup object for the worksheet
Aspose.Cells.PageSetup pageSetup = worksheet.PageSetup;

// Set FitToPagesWide and FitToPagesTall to 0 for unlimited width and height
pageSetup.FitToPagesWide = 0;
pageSetup.FitToPagesTall = 0;

// Save the workbook to a file
workbook.Save("Result.xlsx");
