// Title: Remove printer‑related page setup settings from all worksheets in an XLSX file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet, clears the PrintArea, PrintTitleRows, PrintTitleColumns, and resets FitToPagesWide/FitToPagesTall, then saves the workbook as a new XLSX file. | Show how to use the PageSetup object in Aspose.Cells to strip all printer configuration from each sheet and produce a clean Excel workbook.
// Common Searches: aspnet remove print area and titles from all sheets in an existing Excel workbook | c# aspose.cells clear page setup settings for every worksheet | how to reset fit-to-page options in Aspose.Cells before saving workbook | strip printer settings from XLSX using Aspose.Cells .NET API
// Tags: aspocells strip worksheet page setup | remove printer configuration from Excel using Aspose.Cells | clean Excel workbook by clearing page setup | aspocells reset worksheet print settings | save cleaned XLSX with Aspose.Cells

using Aspose.Cells;

// The program loads 'input.xlsx', iterates over each worksheet, clears print‑related PageSetup properties (PrintArea, PrintTitleRows, PrintTitleColumns, FitToPagesWide/Tall), and saves the cleaned workbook as 'cleaned.xlsx'.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Remove printer‑related settings from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Clear the defined print area
            pageSetup.PrintArea = "";

            // Clear rows and columns set as print titles
            pageSetup.PrintTitleRows = "";
            pageSetup.PrintTitleColumns = "";

            // Reset any fit‑to‑page scaling
            pageSetup.FitToPagesWide = 0;
            pageSetup.FitToPagesTall = 0;
        }

        // Save the cleaned workbook as a new XLSX file
        workbook.Save("cleaned.xlsx", SaveFormat.Xlsx);
    }
}
