using System;
using System.Drawing;
using Aspose.Cells;

class SetWorksheetTabColor
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet tab color to blue
        worksheet.TabColor = Color.Blue;

        // Save the workbook (lifecycle save)
        workbook.Save("TabColorBlue.xlsx", SaveFormat.Xlsx);
    }
}