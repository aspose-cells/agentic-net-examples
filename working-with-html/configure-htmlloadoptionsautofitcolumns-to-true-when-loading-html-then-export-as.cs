using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        // Create HTML load options (default settings)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Load the HTML file into a workbook using the configured options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Export the workbook to XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}