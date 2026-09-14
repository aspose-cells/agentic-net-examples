// Title: Set worksheet page setup to fit 1 page wide and 2 pages tall (single‑column layout) with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to configure a worksheet's PageSetup so that the printed output fits 1 page wide and 2 pages tall, then save the workbook as an .xlsx file. | Show how to apply a single‑column print layout by setting FitToPagesWide = 1 and FitToPagesTall = 2 on a worksheet with Aspose.Cells in C#. | Create a reusable C# method that accepts a Workbook and sets its first worksheet to a 1‑by‑2 page fit using Aspose.Cells PageSetup properties.
// Common Searches: Aspose.Cells C# set worksheet FitToPagesWide to 1 and FitToPagesTall to 2 | How to configure single column print layout in Aspose.Cells .NET | Fit worksheet to specific number of pages using Aspose.Cells PageSetup | C# Aspose.Cells page scaling fit to 1 page wide 2 pages tall example | Set page setup for workbook to print on 1x2 pages with Aspose.Cells
// Tags: Aspose.Cells worksheet page setup fit-to-page | C# page setup scaling properties Aspose.Cells | worksheet print scaling Aspose.Cells | Aspose.Cells page scaling configuration | save workbook as xlsx Aspose.Cells

using Aspose.Cells;

// // Configures the first worksheet's PageSetup to fit 1 page wide by 2 pages tall (single‑column layout) and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page setup for a single‑column layout
        sheet.PageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
        sheet.PageSetup.FitToPagesTall = 2;   // Fit to 2 pages tall

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
