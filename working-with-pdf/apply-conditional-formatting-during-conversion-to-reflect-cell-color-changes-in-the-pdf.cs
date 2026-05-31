using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConditionalFormattingPdfDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample numeric data in column A
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(30);
        worksheet.Cells["A3"].PutValue(60);
        worksheet.Cells["A4"].PutValue(90);

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A4 for the conditional formatting
        CellArea area = new CellArea { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 0 };
        fcc.AddArea(area);

        // Condition 1: values between 20 and 50 → Yellow background
        int condition1 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "20", "50");
        FormatCondition fc1 = fcc[condition1];
        fc1.Style.BackgroundColor = Color.Yellow;

        // Condition 2: values greater than 50 → LightGreen background and bold font
        int condition2 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
        FormatCondition fc2 = fcc[condition2];
        fc2.Style.BackgroundColor = Color.LightGreen;
        fc2.Style.Font.IsBold = true;

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Merge conditional formatting areas before rendering to ensure colors appear in PDF
            MergeAreas = true,
            // Calculate any formulas before rendering (optional for this example)
            CalculateFormula = true
        };

        // Save the workbook as a PDF; conditional formatting colors will be reflected
        workbook.Save("ConditionalFormattingDemo.pdf", pdfOptions);
    }
}