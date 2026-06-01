using System;
using Aspose.Cells;

namespace AsposeCellsUnionRangeNumberFormat
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a union range that includes columns M and O rows 1 to 5
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("M1:M5,O1:O5", 0);

            // Define a style with a uniform number format (e.g., two decimal places)
            Style numberStyle = workbook.CreateStyle();
            numberStyle.Custom = "0.00"; // Uniform number format

            // Apply the style to the entire union range
            StyleFlag flag = new StyleFlag { All = true };
            unionRange.ApplyStyle(numberStyle, flag);

            // Save the workbook
            workbook.Save("UnionRangeNumberFormat.xlsx");
        }
    }
}