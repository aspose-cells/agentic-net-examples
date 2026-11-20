// Title: How to make the first row of an Excel worksheet bold using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set first row font to bold in Excel workbook | apply bold style to header range A1:D1 using Aspose.Cells .NET | Aspose.Cells create bold header row programmatically in C# | using StyleFlag to make header text bold with Aspose.Cells | populate Excel header and apply bold formatting via Aspose.Cells API
// Tags: apply bold font Aspose.Cells C# | header cells formatting Aspose.Cells | font bold attribute style flag Aspose.Cells | create style range Excel .NET | save workbook as Output.xlsx Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBoldExample
{
    // The example creates a new workbook, writes header values to the first row, defines a bold font style with a StyleFlag, applies the style to the header range (A1:D1), and saves the file as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define header values
                string[] headers = { "ID", "Name", "Date", "Amount" };

                // Populate header cells in the first row (A1:D1)
                for (int col = 0; col < headers.Length; col++)
                {
                    sheet.Cells[0, col].PutValue(headers[col]);
                }

                // Create a style with bold font
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Define a style flag to apply only the font bold attribute
                StyleFlag styleFlag = new StyleFlag
                {
                    FontBold = true
                };

                // Apply the bold style to the header range (A1:D1)
                AsposeRange headerRange = sheet.Cells.CreateRange(0, 0, 1, headers.Length);
                headerRange.ApplyStyle(boldStyle, styleFlag);

                // Save the workbook
                workbook.Save("Output.xlsx");
                Console.WriteLine("Workbook saved successfully as Output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
