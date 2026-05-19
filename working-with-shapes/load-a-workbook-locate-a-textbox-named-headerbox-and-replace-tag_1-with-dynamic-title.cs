using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Dynamic title to replace the placeholder with
        string dynamicTitle = "Quarterly Report";

        // Load the workbook from file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Locate the TextBox named "HeaderBox"
        TextBox headerBox = worksheet.TextBoxes["HeaderBox"];

        // If the TextBox exists and contains text, replace the placeholder
        if (headerBox != null && headerBox.Text != null)
        {
            headerBox.Text = headerBox.Text.Replace("<TAG_1>", dynamicTitle);
        }

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}