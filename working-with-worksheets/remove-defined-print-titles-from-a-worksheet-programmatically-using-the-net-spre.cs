using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleRemoval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a blank workbook

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Example: Define print titles (rows and columns) for demonstration
            // ------------------------------------------------------------
            // Repeat the first row on every printed page
            worksheet.PageSetup.PrintTitleRows = "$1:$1";
            // Repeat the first column on every printed page
            worksheet.PageSetup.PrintTitleColumns = "$A:$A";

            // At this point the worksheet has print titles defined.
            // ------------------------------------------------------------
            // Remove the defined print titles
            // ------------------------------------------------------------
            // Setting the properties to an empty string clears the titles.
            worksheet.PageSetup.PrintTitleRows = string.Empty;
            worksheet.PageSetup.PrintTitleColumns = string.Empty;

            // Verify removal (optional)
            Console.WriteLine("PrintTitleRows after removal: " + 
                (string.IsNullOrEmpty(worksheet.PageSetup.PrintTitleRows) ? "None" : worksheet.PageSetup.PrintTitleRows));
            Console.WriteLine("PrintTitleColumns after removal: " + 
                (string.IsNullOrEmpty(worksheet.PageSetup.PrintTitleColumns) ? "None" : worksheet.PageSetup.PrintTitleColumns));

            // Save the workbook to a file (the file will have no print titles)
            workbook.Save("PrintTitleRemoved.xlsx", SaveFormat.Xlsx);
        }
    }
}