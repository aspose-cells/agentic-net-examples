// Title: Import a mixed‑type ArrayList into an Excel worksheet and set date and currency formats for specific columns using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to import an ArrayList of object arrays into a worksheet starting at A1, then apply a mm‑dd‑yy date format to column C and a $#,##0.00 currency format to column D, and save the file. | Create custom Style objects for date and currency, apply them with StyleFlag to the appropriate column ranges after calling Worksheet.Cells.ImportArrayList in C#. | Handle mixed data types (string, int, DateTime, double, bool) from an ArrayList, import them into a new workbook, format the birth‑date and salary columns, and write the workbook to disk.
// Common Searches: Aspose.Cells C# import ArrayList with mixed data types and format columns | how to apply date number format to a column after ImportArrayList in Aspose.Cells | set currency style for a specific column in an Excel file using Aspose.Cells .NET | example of using Worksheet.Cells.ImportArrayList with custom styles for date and currency | C# Aspose.Cells import object[] rows and format birthdate column
// Tags: ImportArrayList with mixed data Aspose.Cells | apply date number format column Aspose.Cells | apply currency style column Aspose.Cells | custom StyleFlag column formatting .NET | C# worksheet data import and styling Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

// Demonstrates importing a mixed‑type ArrayList into a worksheet, then applying a mm‑dd‑yy date format to the BirthDate column and a $#,##0.00 currency format to the Salary column, and finally saving the workbook as MixedData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare mixed data (each row is an object array)
            ArrayList data = new ArrayList();
            data.Add(new object[] { "John", 30, new DateTime(1990, 5, 1), 1234.56, true });
            data.Add(new object[] { "Alice", 25, new DateTime(1995, 8, 15), 7890.12, false });
            data.Add(new object[] { "Bob", 40, new DateTime(1980, 12, 31), 3456.78, true });

            // Import the data starting at cell A1 (no header row)
            worksheet.Cells.ImportArrayList(data, 0, 0, false);

            // Apply date format to the BirthDate column (index 2)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // mm-dd-yy
            StyleFlag dateFlag = new StyleFlag { NumberFormat = true };
            worksheet.Cells.CreateRange(0, 2, data.Count, 1).ApplyStyle(dateStyle, dateFlag);

            // Apply currency format to the Salary column (index 3)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // $#,##0.00
            StyleFlag currencyFlag = new StyleFlag { NumberFormat = true };
            worksheet.Cells.CreateRange(0, 3, data.Count, 1).ApplyStyle(currencyStyle, currencyFlag);

            // Save the workbook
            string outputPath = "MixedData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
