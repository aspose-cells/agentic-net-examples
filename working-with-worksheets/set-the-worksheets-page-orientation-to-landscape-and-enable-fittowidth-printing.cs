// Title: Configure a worksheet to print in landscape orientation and fit to page width with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets the PageSetup.Orientation of a worksheet to Landscape and configures FitToPagesWide = 1 while leaving FitToPagesTall unrestricted using Aspose.Cells. | Provide a concise Aspose.Cells example that prints an Excel sheet in landscape mode and scales the output to a single page wide.
// Common Searches: Aspose.Cells C# set worksheet print orientation to landscape | how to make Excel sheet fit one page wide using Aspose.Cells | PageSetup FitToPagesWide property example Aspose.Cells .NET | print Excel workbook landscape with fit-to-width scaling Aspose.Cells
// Tags: PageSetup orientation landscape Aspose.Cells | page width scaling Aspose.Cells .NET | worksheet print configuration C# Aspose.Cells | Excel export landscape Aspose.Cells

using Aspose.Cells;

// // Creates a new workbook, sets the first worksheet's page orientation to landscape, configures the page setup to fit the content to one page wide (height unrestricted), and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // {CreateWorkbook}
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Enable fit‑to‑width printing (fit to 1 page wide, unlimited pages tall)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // {SaveWorkbook}
        workbook.Save("output.xlsx");
    }
}
