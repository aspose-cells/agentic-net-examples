// Title: Import a DateTime array into a new worksheet and set short‑date format with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds a worksheet, converts a DateTime[] to object[], imports the dates vertically starting at A1, defines a style with built‑in number format 14 (short date), applies the style via StyleFlag, and saves the file as DateImportShortFormat.xlsx.
// Keywords: Aspose.Cells import DateTime array | C# short date format Excel | ImportObjectArray Aspose.Cells | built‑in number format 14 | apply date style Aspose.Cells .NET | create worksheet Aspose.Cells | Excel date formatting C#
// Common Searches: how to import dates into Aspose.Cells worksheet | set short date number format in Aspose.Cells .NET | ImportObjectArray with DateTime values C# | apply built‑in date style to a range Aspose.Cells | C# Aspose.Cells example for date formatting
// Developer Intent: Load a DateTime[] into a fresh worksheet and display the cells using the short‑date pattern.
// Use Cases: Generate a transaction ledger where dates are imported from a C# list and shown as concise short dates. | Build a project timeline spreadsheet by importing schedule dates and formatting them for end‑user readability. | Export appointment data from an application to Excel with proper short‑date rendering for reporting.
// AI Prompts: Provide C# code that uses Aspose.Cells to import a DateTime[] into a worksheet and apply the built‑in short date format (number 14). | Show how to use ImportObjectArray together with CreateStyle and StyleFlag to format dates as short dates in Aspose.Cells for .NET. | Explain step‑by‑step how to create a workbook, import dates vertically, and set only the number format for the imported range.

using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsDateImportDemo
{
    // C# example that creates a workbook, adds a worksheet, converts a DateTime[] to object[], imports the dates vertically starting at A1, defines a style with built‑in number format 14 (short date), applies the style via StyleFlag, and saves the file as DateImportShortFormat.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare an array of DateTime values
                DateTime[] dateArray = new DateTime[]
                {
                    new DateTime(2023, 1, 15),
                    new DateTime(2023, 2, 20),
                    new DateTime(2023, 3, 25),
                    new DateTime(2023, 4, 10)
                };

                // Convert to object[] because ImportObjectArray expects object[]
                object[] objArray = dateArray.Cast<object>().ToArray();

                // Import the dates vertically starting at cell A1 (row 0, column 0)
                worksheet.Cells.ImportObjectArray(objArray, 0, 0, true);

                // Apply short date format (built‑in number format 14) to the imported range
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Short date pattern

                StyleFlag styleFlag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Create a range that covers the imported dates
                int rowCount = dateArray.Length;
                Aspose.Cells.Range dateRange = worksheet.Cells.CreateRange(0, 0, rowCount, 1);
                dateRange.ApplyStyle(dateStyle, styleFlag);

                // Save the workbook
                string outputPath = "DateImportShortFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
