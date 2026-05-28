using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable the Ribbon UI by clearing the RibbonXml property
        workbook.RibbonXml = string.Empty;

        // Prepare ODS save options (default settings are sufficient)
        OdsSaveOptions odsOptions = new OdsSaveOptions();

        // Save the workbook as an ODS file; the RibbonXml setting will be persisted
        workbook.Save("Result.ods", odsOptions);
    }
}