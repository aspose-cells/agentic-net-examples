// Title: C# – Create a Worksheet, Import a DateTime Array, and Apply a Short Date Format with Aspose.Cells
// Description: Demonstrates how to create a new workbook, import an array of DateTime objects vertically into cells starting at A1, define a custom short‑date style (mm/dd/yyyy), apply the style using a StyleFlag that targets only number formatting, and save the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# import DateTime array | short date format Excel Aspose | ImportObjectArray date style | .NET Excel date formatting | Create worksheet Aspose.Cells | StyleFlag number format
// Common Searches: how to import DateTime array into Aspose.Cells worksheet | apply short date format to a range in Aspose.Cells .NET | Aspose.Cells ImportObjectArray custom number format | C# set mm/dd/yyyy format for Excel cells using Aspose | save workbook after styling dates Aspose.Cells
// Developer Intent: Import a DateTime collection into a new worksheet and format the cells to show a short date pattern.
// Use Cases: Generate a transaction log where all dates must appear uniformly as mm/dd/yyyy. | Export project milestone dates to Excel with consistent short‑date formatting for stakeholder reports. | Create a calendar export that preserves date values and displays them in a standard short format for downstream analysis.
// AI Prompts: Provide C# code that creates an Aspose.Cells workbook, imports a DateTime[] vertically, applies the "mm/dd/yyyy" format to the range, and saves the file. | Explain how StyleFlag can be used to apply only number formatting when styling imported dates in Aspose.Cells. | Show a step‑by‑step example of importing dates, setting a custom short date style, and exporting the workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDateImportDemo
{
    // Demonstrates how to create a new workbook, import an array of DateTime objects vertically into cells starting at A1, define a custom short‑date style (mm/dd/yyyy), apply the style using a StyleFlag that targets only number formatting, and save the file as an Excel workbook using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare an array of DateTime objects
                object[] dateArray = new object[]
                {
                    new DateTime(2023, 1, 15),
                    new DateTime(2023, 2, 20),
                    new DateTime(2023, 3, 25)
                };

                // Import the dates vertically starting at cell A1 (row 0, column 0)
                worksheet.Cells.ImportObjectArray(dateArray, 0, 0, true);

                // Create a style with a short date format (e.g., mm/dd/yyyy)
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "mm/dd/yyyy";

                // Apply the style to the imported range
                int rowsCount = dateArray.Length;
                AsposeRange dateRange = worksheet.Cells.CreateRange(0, 0, rowsCount, 1);
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true; // Apply only number format
                dateRange.ApplyStyle(dateStyle, flag);

                // Define output file path
                string outputPath = "DateArrayImport.xlsx";

                // Save the workbook (overwrite if exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
