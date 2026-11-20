// Title: Create a worksheet, import a DateTime array, and apply the short date format (ID 14) with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# import DateTime array into a worksheet | apply short date number format (ID 14) to cells using Aspose.Cells | how to style a range with built‑in date format in Aspose.Cells .NET | save a workbook containing dates as XLSX with Aspose.Cells C#
// Tags: import DateTime array Aspose.Cells C# | apply short date number format Aspose.Cells | style range with built‑in date format Aspose.Cells | create worksheet and save XLSX Aspose.Cells | use ImportObjectArray for dates Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds a worksheet named "DatesSheet", imports an array of DateTime values into cells starting at A1 using ImportObjectArray, defines a Style with the built‑in short date number format (ID 14), applies the style to the imported range via ApplyStyle, and saves the workbook as DatesWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet named "DatesSheet"
            Worksheet sheet = workbook.Worksheets.Add("DatesSheet");

            // Define an array of DateTime values
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 2, 14),
                new DateTime(2023, 3, 15),
                new DateTime(2023, 4, 1)
            };

            // Convert DateTime array to object array for ImportObjectArray
            object[] dateObjects = Array.ConvertAll(dates, d => (object)d);

            // Import the dates array into the worksheet starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportObjectArray(dateObjects, 0, 0, false);

            // Create a style with the built‑in short date number format (ID 14)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Short date pattern

            // Apply only the number format
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.NumberFormat = true;

            // Apply the style to the range that contains the imported dates
            Aspose.Cells.Range dateRange = sheet.Cells.CreateRange(0, 0, dates.Length, 1);
            dateRange.ApplyStyle(dateStyle, styleFlag);

            // Save the workbook to a file
            string outputPath = "DatesWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
