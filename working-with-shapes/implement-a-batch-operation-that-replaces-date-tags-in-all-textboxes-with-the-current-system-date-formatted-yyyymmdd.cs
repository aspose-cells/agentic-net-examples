using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsTextBoxDateReplace
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Prepare the replacement string: current date in yyyy-MM-dd format
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of TextBox objects on the worksheet
                TextBoxCollection textBoxes = sheet.TextBoxes;

                // Loop through each TextBox
                for (int i = 0; i < textBoxes.Count; i++)
                {
                    TextBox tb = textBoxes[i];

                    // Use FontSettingCollection.Replace(string, string) to replace <DATE> tags
                    // This method is part of the provided API and operates on the TextBody of the shape
                    tb.TextBody.Replace("<DATE>", currentDate);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}