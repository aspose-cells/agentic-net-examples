using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsUnionRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a regular range A1:A5 and assign a name to it
                AsposeRange namedRange = worksheet.Cells.CreateRange("A1:A5");
                namedRange.Name = "MyNamedRange";

                // Populate the named range with sample data
                for (int i = 0; i < namedRange.RowCount; i++)
                {
                    namedRange[i, 0].PutValue($"Item {i + 1}");
                }

                // Create a UnionRange that includes the named range and another address C1:C5
                // The named range can be referenced by its name in the address string
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("MyNamedRange,C1:C5", 0);

                // Define a style with a solid background color
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightGreen;

                // Apply the style to the entire union range
                StyleFlag flag = new StyleFlag { All = true };
                unionRange.ApplyStyle(style, flag);

                // Save the workbook
                workbook.Save("UnionRangeWithNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}