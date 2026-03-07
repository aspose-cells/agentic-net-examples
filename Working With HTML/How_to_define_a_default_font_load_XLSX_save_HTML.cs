using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Define the default font to be used when loading the workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.DefaultStyleSettings.FontName = "Arial";

        // Load the XLSX file with the specified load options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure the workbook's default style also uses the same font (optional)
        workbook.DefaultStyle.Font.Name = "Arial";

        // Set HTML save options and specify the default font for HTML export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Arial";

        // Save the workbook as HTML using the defined options
        workbook.Save("output.html", htmlOptions);
    }
}