using Aspose.Cells;

class SetPageOptionsDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = sheet.PageSetup;

        // Set margins (centimeters)
        pageSetup.TopMargin = 1.5;
        pageSetup.BottomMargin = 1.5;
        pageSetup.LeftMargin = 2.0;
        pageSetup.RightMargin = 2.0;

        // Set page orientation to Landscape
        pageSetup.Orientation = PageOrientationType.Landscape;

        // Set paper size to A4
        pageSetup.PaperSize = PaperSizeType.PaperA4;

        // Center the sheet horizontally and vertically when printed
        pageSetup.CenterHorizontally = true;
        pageSetup.CenterVertically = true;

        // Configure print options
        pageSetup.PrintGridlines = true;      // Print cell gridlines
        pageSetup.PrintHeadings = true;       // Print row/column headings
        pageSetup.PrintDraft = false;         // Print with graphics
        pageSetup.BlackAndWhite = false;      // Print in color

        // Fit the worksheet to 1 page wide and 1 page tall (no percent scaling)
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 1;
        pageSetup.IsPercentScale = false;

        // Set the first page number and disable automatic numbering
        pageSetup.FirstPageNumber = 5;
        pageSetup.IsAutoFirstPageNumber = false;

        // Define rows and columns to repeat on each printed page
        pageSetup.PrintTitleRows = "$1:$1";    // Repeat first row
        pageSetup.PrintTitleColumns = "$A:$A"; // Repeat first column

        // Save the workbook (lifecycle save)
        workbook.Save("PageOptionsDemo.xlsx");
    }
}