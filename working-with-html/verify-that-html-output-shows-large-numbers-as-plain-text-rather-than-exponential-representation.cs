using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Large numeric value that Excel would normally display in scientific notation
        double largeNumber = 1234567890123456.0;
        worksheet.Cells["A1"].PutValue(largeNumber);

        // Configure Aspose.Cells to keep full precision when converting to HTML.
        // Setting SignificantDigitsType to Digits15 forces plain decimal output.
        CellsHelper.SignificantDigitsType = SignificantDigitsType.Digits15;

        // Prepare HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        // Ensure the whole cell content is exported regardless of column width
        saveOptions.FormatDataIgnoreColumnWidth = true;

        // Save the workbook as HTML
        string htmlFile = "LargeNumber.html";
        workbook.Save(htmlFile, saveOptions);

        // Load the generated HTML and verify that the number appears as plain text
        string htmlContent = File.ReadAllText(htmlFile);
        // The expected plain representation (no exponent) of the large number
        string expectedPlain = largeNumber.ToString("F0");
        bool isPlain = htmlContent.Contains(expectedPlain);

        Console.WriteLine($"HTML contains plain number \"{expectedPlain}\": {isPlain}");
    }
}