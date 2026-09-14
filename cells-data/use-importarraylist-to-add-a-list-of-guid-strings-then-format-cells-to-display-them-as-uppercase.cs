// Title: Import uppercase GUID strings from an ArrayList into an Excel worksheet and apply a monospaced font using Aspose.Cells for .NET
// AI Prompts: Generate five GUID values, convert them to uppercase strings, add them to an ArrayList, and call Cells.ImportArrayList to write them vertically starting at cell A1. | Create a Style object with the Consolas font, enable the FontName flag, and apply the style to the range that contains the imported GUIDs. | Save the workbook as GuidUppercase.xlsx after importing the GUID list and applying the custom font style.
// Common Searches: Aspose.Cells C# import ArrayList of GUIDs vertically into a worksheet | How to set a monospaced font for a range after using ImportArrayList in Aspose.Cells | Create uppercase GUID strings and write them to Excel with Aspose.Cells .NET | Apply custom style to imported cells in an Aspose.Cells workbook
// Tags: vertical GUID import Aspose.Cells | apply Consolas monospaced font to Excel range | uppercase GUID strings in .NET workbook | style imported cells with custom font Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsGuidExample
{
    // The program creates a new workbook, generates five uppercase GUID strings, stores them in an ArrayList, imports them vertically into the first worksheet starting at A1 using Cells.ImportArrayList, applies a Consolas monospaced font style to the imported range, and saves the file as GuidUppercase.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList of GUID strings in uppercase
            ArrayList guidList = new ArrayList();
            for (int i = 0; i < 5; i++)
            {
                // Generate a new GUID, convert to string and make it uppercase
                guidList.Add(Guid.NewGuid().ToString().ToUpper());
            }

            // Import the GUID list vertically starting at cell A1 (row 0, column 0)
            // Parameters: (ArrayList, firstRow, firstColumn, isVertical)
            cells.ImportArrayList(guidList, 0, 0, true);

            // Optionally, apply a style to the imported range (e.g., set the font to a monospaced type)
            Style style = workbook.CreateStyle();
            style.Font.Name = "Consolas";
            StyleFlag flag = new StyleFlag();
            flag.FontName = true;

            // Apply the style to the range that contains the GUIDs
            int lastRow = guidList.Count - 1;
            cells.CreateRange(0, 0, lastRow + 1, 1).ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("GuidUppercase.xlsx");
        }
    }
}
