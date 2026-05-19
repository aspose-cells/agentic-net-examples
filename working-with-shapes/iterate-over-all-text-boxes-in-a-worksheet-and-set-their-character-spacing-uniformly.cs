using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add sample text boxes to demonstrate the operation
        int tbIndex1 = worksheet.TextBoxes.Add(1, 1, 100, 50);
        worksheet.TextBoxes[tbIndex1].Text = "First TextBox";

        int tbIndex2 = worksheet.TextBoxes.Add(2, 2, 150, 60);
        worksheet.TextBoxes[tbIndex2].Text = "Second TextBox";

        // Define the uniform character spacing value
        double uniformSpacing = 2.0; // Adjust as needed

        // Iterate over all TextBox objects in the worksheet and set their spacing
        foreach (TextBox textBox in worksheet.TextBoxes)
        {
            textBox.TextOptions.Spacing = uniformSpacing;
        }

        // Save the workbook with the updated text boxes
        workbook.Save("TextBoxesUniformSpacing.xlsx");
    }
}