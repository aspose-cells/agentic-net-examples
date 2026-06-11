using System;
using System.Drawing;
using Aspose.Cells;

class SetWorksheetTabColor
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet tab color to teal
        worksheet.TabColor = Color.Teal;

        // Save the workbook (lifecycle: save)
        workbook.Save("WorksheetTabColor_Teal.xlsx");
    }
}