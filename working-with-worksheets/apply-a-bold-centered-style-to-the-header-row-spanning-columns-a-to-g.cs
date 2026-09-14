// Title: Apply a bold, centered style to the header row spanning columns A‑G in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a bold, center‑aligned style with Aspose.Cells and applies it to cells A1 through G1 of a new workbook. | Generate a program that defines a StyleFlag for FontBold and HorizontalAlignment, applies the style to the first row across columns A‑G, and saves the file as HeaderStyled.xlsx.
// Common Searches: aspocells C# set bold and center alignment for header row A1:G1 | how to apply a style to a range of columns in Aspose.Cells .NET | C# Aspose.Cells create StyleFlag for FontBold and HorizontalAlignment | apply formatting to first row across multiple columns using Aspose.Cells workbook | Aspose.Cells example for styling header row in Excel file
// Tags: Aspose.Cells apply style to range | C# bold centered header row Aspose.Cells | StyleFlag FontBold HorizontalAlignment Aspose.Cells | CreateRange A1 G1 Aspose.Cells | save workbook HeaderStyled.xlsx Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderStyle
{
    // The program creates a new workbook, defines a bold and center‑aligned style, applies it to the header row covering columns A‑G via a range, and saves the result as HeaderStyled.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style: bold font and centered alignment
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.HorizontalAlignment = TextAlignmentType.Center;

                // Specify which style properties to apply
                StyleFlag styleFlag = new StyleFlag
                {
                    FontBold = true,
                    HorizontalAlignment = true
                };

                // Define the range covering columns A to G in the first row (row index 0)
                // Parameters: startRow, startColumn, totalRows, totalColumns
                AsposeRange headerRange = worksheet.Cells.CreateRange(0, 0, 1, 7);
                headerRange.ApplyStyle(headerStyle, styleFlag);

                // Save the workbook
                workbook.Save("HeaderStyled.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
