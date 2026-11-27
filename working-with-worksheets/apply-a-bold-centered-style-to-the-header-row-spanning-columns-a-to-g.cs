// Title: How to apply bold and centered formatting to the header row (A1:G1) using Aspose.Cells for .NET
// Common Searches: c# aspnet apply bold font and center alignment to Excel header row using Aspose.Cells | aspocells styleflag to format range A1:G1 in a workbook | set horizontal alignment to center and make text bold in Excel with Aspose.Cells C# | apply custom style to the first row of an XLSX file using Aspose.Cells .NET API | create and apply a style object for header cells in an Aspose.Cells workbook
// Tags: bold font style Aspose.Cells C# | horizontal center alignment Aspose.Cells | styleflag usage for range formatting Aspose.Cells | header row Excel styling Aspose.Cells | apply style to cell range Aspose.Cells

using System;
using Aspose.Cells;

namespace HeaderStyleExample
{
    // The example creates a new XLSX workbook, writes header text into cells A1 through G1, defines a style with a bold font and horizontal center alignment, configures a StyleFlag to apply those attributes, applies the style to the A1:G1 range, and saves the workbook as HeaderStyled.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Optionally put some sample header text in the cells A1 to G1
                string[] headers = { "Column A", "Column B", "Column C", "Column D", "Column E", "Column F", "Column G" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[0, i].PutValue(headers[i]); // Row 0 is the first row (A1, B1, ...)
                }

                // Create a style object
                Style headerStyle = workbook.CreateStyle();

                // Set the font to bold
                headerStyle.Font.IsBold = true;

                // Center the text horizontally
                headerStyle.HorizontalAlignment = TextAlignmentType.Center;

                // Define which style attributes should be applied
                StyleFlag flag = new StyleFlag
                {
                    FontBold = true,
                    HorizontalAlignment = true
                };

                // Apply the style to the range A1:G1
                Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1:G1");
                headerRange.ApplyStyle(headerStyle, flag);

                // Save the workbook to a file
                string outputPath = "HeaderStyled.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'. Header row styled as bold and centered.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
