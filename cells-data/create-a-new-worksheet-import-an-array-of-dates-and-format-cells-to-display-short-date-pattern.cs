// Title: Create a worksheet, import a DateTime array vertically, and format cells with Excel short date pattern using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a new workbook, adds a worksheet, imports a DateTime[] into column A starting at A1, and applies the built‑in short date number format (format 14) to the imported range. | Show how to apply a number‑format style to a specific cell range after importing data with Aspose.Cells in C#. | Adapt the example to import the dates horizontally across row 1 and use a custom date format such as "dd-MMM-yyyy" instead of the built‑in short date.
// Common Searches: Aspose.Cells C# import DateTime array into worksheet and set short date format | How to apply Excel built‑in short date (format 14) to a range using Aspose.Cells | Import dates vertically starting at A1 with Aspose.Cells .NET | Set number format for imported dates in Aspose.Cells workbook | C# Aspose.Cells example for date array import and formatting
// Tags: date array import Aspose.Cells | short date number format Aspose.Cells | worksheet creation Aspose.Cells | vertical date import Aspose.Cells | Excel format 14 Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsDateImportDemo
{
    // // This program creates a new workbook, accesses the first worksheet, imports an array of DateTime values vertically starting at A1, applies the built‑in short date number format (format 14) to the imported cells, and saves the workbook as ImportedDatesShortDate.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare an array of DateTime values to import
            object[] dateArray = new object[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 2, 14),
                new DateTime(2023, 3, 31),
                new DateTime(2023, 4, 15),
                new DateTime(2023, 5, 20)
            };

            // Import the dates vertically starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportObjectArray(dateArray, 0, 0, true);

            // Apply short date format to the imported cells
            int rowsCount = dateArray.Length;
            // Create a range that covers the imported dates
            var dateRange = worksheet.Cells.CreateRange(0, 0, rowsCount, 1);
            // Create a style with short date number format (Excel built‑in format 14)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Short date pattern (e.g., m/d/yyyy)
            // Specify that only the number format should be applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            // Apply the style to the range
            dateRange.ApplyStyle(dateStyle, flag);

            // Save the workbook
            workbook.Save("ImportedDatesShortDate.xlsx");
        }
    }
}
