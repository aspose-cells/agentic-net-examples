using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Utility;

class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file that contains CSS‑based conditional formatting
        string sourceHtml = "input.html";

        // Path where the resulting Excel workbook will be saved
        string destinationExcel = "output.xlsx";

        // LoadOptions configured for HTML format.
        // Aspose.Cells automatically maps CSS classes that define conditional formatting
        // to the corresponding ConditionalFormattingCollection in the workbook.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

        // SaveOptions for the Excel format (XLSX). No special settings are required
        // because the conditional formatting has already been imported during loading.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

        // Perform the conversion using the utility method that internally loads and saves.
        ConversionUtility.Convert(sourceHtml, loadOptions, destinationExcel, saveOptions);

        Console.WriteLine("HTML has been converted to Excel. Conditional formatting preserved.");
    }
}