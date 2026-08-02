using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Disable the Ribbon UI by clearing the RibbonXml property
        workbook.RibbonXml = string.Empty;

        // Prepare ODS save options (optional settings can be adjusted here)
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        odsOptions.RefreshChartCache = true; // ensure chart data is refreshed

        // Save the workbook as an ODS file with the specified options
        workbook.Save("output.ods", odsOptions);
    }
}