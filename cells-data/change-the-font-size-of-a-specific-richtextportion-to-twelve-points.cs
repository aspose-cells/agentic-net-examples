// Title: How to set a 12‑point font size for a specific RichTextPortion in an Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Font.Size property of a RichTextPortion that represents characters 6‑11 in cell A1 to 12 points with Aspose.Cells. | Apply a 12‑point font size to a selected substring inside an Excel cell using the Characters method in C#. | Change the font size of a partial cell string to 12 points via the FontSetting object in Aspose.Cells.
// Common Searches: Aspose.Cells C# change font size of part of a cell text | Set font size for specific characters in Excel using Aspose.Cells .NET | How to use Characters method to format a substring in an Excel worksheet with Aspose
// Tags: Aspose.Cells RichTextPortion font size | C# Characters method partial text formatting | Excel cell substring FontSetting usage | Aspose.Cells set partial text style .xlsx | Font.Size property Aspose.Cells example

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRichTextPortionDemo
{
    // Creates a workbook, writes "Hello Aspose" to cell A1, selects the substring "Aspose" via cell.Characters, sets its Font.Size to 12 points, and saves the file as RichTextPortionFontSize.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the cell value that will contain rich text
            Cell cell = worksheet.Cells["A1"];
            cell.Value = "Hello Aspose";

            // Define the portion of text to modify (e.g., characters 6 to 12 -> "Aspose")
            // Characters(startIndex, length) returns a FontSetting object
            FontSetting richTextPortion = cell.Characters(6, 6);

            // Change the font size of this specific portion to 12 points
            richTextPortion.Font.Size = 12;

            // Optionally, you can also change other font attributes, e.g., make it bold
            // richTextPortion.Font.IsBold = true;

            // Save the workbook to a file
            workbook.Save("RichTextPortionFontSize.xlsx");
        }
    }
}
