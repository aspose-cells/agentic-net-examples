using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access the cells collection
        Cells cells = sheet.Cells;

        // Define the source range and apply a sample style
        AsposeRange sourceRange = cells.CreateRange("A1:C3");
        Style sampleStyle = workbook.CreateStyle();
        sampleStyle.Font.Name = "Arial";
        sampleStyle.Font.Size = 12;
        sampleStyle.Font.IsBold = true;
        sampleStyle.ForegroundColor = Color.Yellow;
        sampleStyle.Pattern = BackgroundType.Solid;
        sourceRange.SetStyle(sampleStyle);

        // Define the destination range
        AsposeRange destinationRange = cells.CreateRange("E1:G3");

        // Copy formatting from the source range to the destination range
        destinationRange.CopyStyle(sourceRange);

        // Save the workbook to an XLSX file
        workbook.Save("CopyFormattingDemo.xlsx");
    }
}