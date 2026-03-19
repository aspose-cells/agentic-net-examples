using System;
using Aspose.Cells;

namespace AsposeCellsRelatedArticleDemo
{
    class Program
    {
        static void Main()
        {
            // URL of the related article that explains loading options for XLSX workbooks
            // (Replace with the actual article link as needed)
            const string relatedArticleUrl = "https://docs.aspose.com/cells/net/loading-workbooks/";

            // Create LoadOptions to control how the workbook is loaded.
            // Here we explicitly set the format to XLSX and enable formula parsing on open.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingFormulaOnOpen = true
            };

            // Load the workbook using the constructor that accepts a file path and LoadOptions.
            // This demonstrates programmatic loading while referencing the related article above.
            Workbook workbook = new Workbook("Sample.xlsx", loadOptions);

            // Example usage: access the first worksheet and print its name.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"Loaded worksheet: {sheet.Name}");

            // Dispose the workbook when done.
            workbook.Dispose();
        }
    }
}