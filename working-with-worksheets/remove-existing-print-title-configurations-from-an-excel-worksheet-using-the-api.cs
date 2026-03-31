using System;
using Aspose.Cells;

class RemovePrintTitleConfigurations
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet from which print titles should be removed
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Clear the rows and columns set as print titles
        pageSetup.PrintTitleRows = "";
        pageSetup.PrintTitleColumns = "";

        // Optionally clear any header/footer settings as part of resetting page‑setup
        pageSetup.ClearHeaderFooter();

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}