using System;
using Aspose.Cells;

namespace AsposeCellsNameUpdate
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a defined name "ReportData"
            Workbook workbook = new Workbook("input.xlsx");

            // Retrieve the Name object with the name "ReportData"
            // The Names collection can be accessed by index or by name string
            Name reportName = workbook.Worksheets.Names["ReportData"];

            // Ensure the name exists before attempting to modify it
            if (reportName != null)
            {
                // Change the Text property to the new name "SummaryData"
                reportName.Text = "SummaryData";
            }
            else
            {
                Console.WriteLine("The defined name 'ReportData' was not found in the workbook.");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}