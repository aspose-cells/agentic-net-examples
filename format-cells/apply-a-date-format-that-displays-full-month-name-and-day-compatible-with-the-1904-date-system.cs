// Title: Aspose.Cells for .NET: Format a Cell as Full Month Name and Day Using the 1904 Date System
// Description: Demonstrates how to enable the 1904 date system in an Aspose.Cells workbook, convert a .NET DateTime to the correct Excel serial value, and apply the custom number format "mmmm d" so the cell shows the full month name and day (e.g., "July 15"). The workbook is saved as FullMonthNameDay_1904.xlsx.
// Keywords: Aspose.Cells C# | custom date format mmmm d | 1904 date system | Excel serial number conversion | full month name day format | legacy Mac Excel compatibility | date formatting .NET | CellsHelper GetDoubleFromDateTime | Excel workbook formatting | Aspose.Cells examples
// Common Searches: Aspose.Cells 1904 date system example | C# format Excel cell month name day | how to use custom date format mmmm d in Aspose.Cells | convert .NET DateTime to Excel serial number 1904 | display full month name in Excel using Aspose.Cells
// Developer Intent: Create an Excel file that uses the 1904 date system and displays dates as "MonthName Day" via a custom format.
// Use Cases: Generating reports for older Mac Excel files that require the 1904 date system. | Converting .NET DateTime values to Excel serial numbers while preserving legacy date calculations. | Applying a consistent "MonthName Day" display across multiple worksheets in automated spreadsheet generation.
// AI Prompts: Show me C# code to enable the 1904 date system in Aspose.Cells and format a cell with "mmmm d". | Explain how to convert a .NET DateTime to an Excel serial number for the 1904 date system using Aspose.Cells. | Provide a complete Aspose.Cells example that saves a workbook with a cell displaying "July 15" in the 1904 date system.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable the 1904 date system in an Aspose.Cells workbook, convert a .NET DateTime to the correct Excel serial value, and apply the custom number format "mmmm d" so the cell shows the full month name and day (e.g., "July 15"). The workbook is saved as FullMonthNameDay_1904.xlsx.
    public class FullMonthNameDayDateFormatDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Enable the 1904 date system
            wb.Settings.Date1904 = true;

            // Get the first worksheet and a target cell
            Worksheet sheet = wb.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Define a DateTime value (July 15, 2023)
            DateTime dateValue = new DateTime(2023, 7, 15);

            // Convert the DateTime to Excel serial number using the 1904 system
            double excelSerial = CellsHelper.GetDoubleFromDateTime(dateValue, wb.Settings.Date1904);

            // Put the serial value into the cell
            cell.PutValue(excelSerial);

            // Apply a custom number format that shows full month name and day (e.g., "July 15")
            Style style = cell.GetStyle();
            style.Custom = "mmmm d";
            cell.SetStyle(style);

            // Determine output file path
            string outputFile = "FullMonthNameDay_1904.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputFile);
        }
    }
}
