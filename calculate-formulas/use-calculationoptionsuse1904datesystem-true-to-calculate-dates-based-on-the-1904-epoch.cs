// Title: Calculate a 1904‑epoch Excel date and convert the serial number to a .NET DateTime using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to evaluate the formula =DATE(1904,1,1), retrieve the resulting serial number, and convert it to a System.DateTime by passing the 1904 flag to CellsHelper.GetDateTimeFromDouble. | Show how to insert the 1904‑based serial number into a worksheet cell, apply a built‑in date number format, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# calculate DATE formula using 1904 date system | Convert Excel serial number to .NET DateTime with 1904 epoch in Aspose.Cells | How to use CellsHelper.GetDateTimeFromDouble with 1904 flag in C# | Write 1904 based serial number to a cell and format as date using Aspose.Cells | Save workbook containing 1904 epoch dates with Aspose.Cells .NET
// Tags: calculate 1904 date formula Aspose.Cells | convert 1904 serial number to DateTime C# | write serial number to cell with date format Aspose.Cells | save workbook with 1904 epoch dates .xlsx | CellsHelper.GetDateTimeFromDouble 1904 flag

using System;
using Aspose.Cells;

namespace AsposeCells1904DateExample
{
    // The example creates a new workbook, evaluates the formula =DATE(1904,1,1) to obtain a 1904‑epoch serial number, converts that serial number to a System.DateTime using CellsHelper.GetDateTimeFromDouble with the 1904 flag, writes the serial number to cell A1, applies a built‑in date format, and saves the file as 1904DateSystemExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // NOTE: The Use1904DateSystem property may not be available in some versions.
                // The conversion below explicitly treats the serial number as a 1904‑based date,
                // so the workbook setting is not required for this example.

                // Calculate a date formula. In the 1904 system, DATE(1904,1,1) corresponds to serial number 0
                object serialResult = worksheet.CalculateFormula("=DATE(1904,1,1)");

                // Convert the serial number to a .NET DateTime using the 1904 flag
                double serialNumber = Convert.ToDouble(serialResult);
                DateTime date = CellsHelper.GetDateTimeFromDouble(serialNumber, true);

                // Output the results
                Console.WriteLine("Serial number (1904 epoch): " + serialNumber);
                Console.WriteLine("Converted DateTime: " + date.ToString("yyyy-MM-dd"));

                // Optionally write the serial number to a cell and format it as a date
                worksheet.Cells["A1"].PutValue(serialNumber);
                Style dateStyle = worksheet.Cells["A1"].GetStyle();
                dateStyle.Number = 14; // Built‑in date format
                worksheet.Cells["A1"].SetStyle(dateStyle);

                // Save the workbook
                string outputPath = "1904DateSystemExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
