// Title: C# – Retrieve the page‑scale percentage when fitting an Aspose.Cells worksheet to a single printed sheet
// Description: Creates a workbook, fills cells A1:J50, sets the print area, applies SetFitToPages(1,1) to force a one‑page layout, builds ImageOrPrintOptions, instantiates SheetRender, reads the PageScale property, prints the scaling percentage, and optionally saves the file.
// Keywords: Aspose.Cells | C# | .NET | PageScale | SheetRender | SetFitToPages | fit-to-page | print scaling | worksheet scaling factor | Excel print layout | calculate page scale
// Common Searches: Aspose.Cells get page scale after SetFitToPages | C# calculate print scaling for Excel worksheet | How to retrieve scaling factor for fit‑to‑one‑page in Aspose.Cells | SheetRender PageScale example | Determine print zoom percentage with Aspose.Cells .NET
// Developer Intent: Find out how to obtain the exact scaling percentage that Aspose.Cells applies when a worksheet is configured to fit on a single printed page.
// Use Cases: Display the calculated scale to users before printing so they know the reduction level. | Adjust margins, headers, or other layout settings based on the retrieved PageScale value. | Log or report the scaling factor when generating batch prints to ensure consistent output. | Synchronize a custom viewer’s zoom level with the actual print scaling.
// AI Prompts: Show code that changes the fit‑to‑page setting to 1 page wide and multiple pages tall while still returning the PageScale value. | Provide a loop that iterates over several print areas, captures each PageScale, and outputs a summary table of scaling percentages. | Explain how to use the PageScale value to set a custom zoom level in a WinForms or WPF worksheet viewer.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, fills cells A1:J50, sets the print area, applies SetFitToPages(1,1) to force a one‑page layout, builds ImageOrPrintOptions, instantiates SheetRender, reads the PageScale property, prints the scaling percentage, and optionally saves the file.
class FitToOnePageDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the sheet with sample data
        for (int i = 0; i < 50; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Define the print area that includes the data
        sheet.PageSetup.PrintArea = "A1:J50";

        // Fit the worksheet to a single page (wide and tall)
        sheet.PageSetup.SetFitToPages(1, 1); // uses PageSetup.SetFitToPages

        // Create rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        // Create a SheetRender after page‑setup changes
        SheetRender render = new SheetRender(sheet, options); // uses SheetRender constructor

        // Retrieve the calculated page scale
        double pageScale = render.PageScale; // uses SheetRender.PageScale

        // Output the scale as a percentage
        Console.WriteLine($"Calculated page scale to fit on one sheet: {pageScale * 100:0.##}%");

        // Save the workbook (optional, demonstrates saving)
        workbook.Save("FitToOnePage.xlsx");
    }
}
