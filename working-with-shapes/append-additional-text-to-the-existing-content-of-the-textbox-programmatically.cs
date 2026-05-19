using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsAppendTextDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet (topRow, leftColumn, height, width)
            int textboxIndex = worksheet.TextBoxes.Add(2, 1, 50, 200);
            TextBox textBox = worksheet.TextBoxes[textboxIndex];

            // Set initial text for the textbox
            textBox.Text = "Hello, ";

            // Append additional text using FontSettingCollection.AppendText
            // TextBody returns a FontSettingCollection instance
            textBox.TextBody.AppendText("World!");

            // Optionally display the final text in console
            Console.WriteLine("Final textbox text: " + textBox.TextBody.Text);

            // Save the workbook
            workbook.Save("AppendTextToTextbox.xlsx");
        }
    }
}