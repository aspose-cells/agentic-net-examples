using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample numeric data
        for (int row = 0; row <= 10; row++)
        {
            worksheet.Cells[row, 0].PutValue(row);
        }

        // Add a 3‑color scale conditional formatting to the range A1:A11
        int cfIndex = worksheet.ConditionalFormattings.Add();
        var cfCollection = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for the conditional formatting
        var area = new CellArea { StartRow = 0, EndRow = 10, StartColumn = 0, EndColumn = 0 };
        cfCollection.AddArea(area);

        // Create the color‑scale condition
        int conditionIndex = cfCollection.AddCondition(FormatConditionType.ColorScale);
        var condition = cfCollection[conditionIndex];

        // Configure the 3‑color scale (Red → Yellow → Green)
        condition.ColorScale.Is3ColorScale = true;
        condition.ColorScale.MinColor = Color.Red;
        condition.ColorScale.MidColor = Color.Yellow;
        condition.ColorScale.MaxColor = Color.Green;

        // Configure PDF save options to preserve conditional formatting colors
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Merges conditional‑formatting areas before rendering, ensuring colors are kept
            MergeAreas = true
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("ConditionalFormattingColors.pdf", pdfOptions);
    }
}