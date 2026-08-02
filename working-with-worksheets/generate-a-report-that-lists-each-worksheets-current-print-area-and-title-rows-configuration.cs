// Title: C# Aspose.Cells: Generate a Print Setup Report of Worksheet Print Areas and Title Rows
// Description: Creates a workbook, assigns sample PrintArea and PrintTitleRows values, adds a "PrintSetupReport" sheet, lists each worksheet’s name with its PrintArea and PrintTitleRows (or "Not Set"), and saves the workbook as PrintSetupReport.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PrintArea | PrintTitleRows | PageSetup | worksheet print settings | Excel report generation | list worksheet print configuration | export print setup to Excel
// Common Searches: Aspose.Cells get PrintArea for each worksheet | C# list PrintTitleRows using Aspose.Cells | generate Excel report of worksheet print settings .NET | how to enumerate PageSetup properties with Aspose.Cells | create summary sheet of print areas in Aspose.Cells
// Developer Intent: Produce an Excel summary that enumerates the print area and title‑row settings for every worksheet in a workbook.
// Use Cases: Audit print configurations across all sheets before bulk printing. | Provide a quick reference sheet for end‑users to see printable ranges and header rows. | Validate that automatically generated reports contain required title rows.
// AI Prompts: Write C# code with Aspose.Cells that adds a report worksheet showing each sheet’s PrintArea and PrintTitleRows and saves the file. | Extend the sample to also capture PrintTitleColumns and apply bold formatting to the header row. | Explain how to handle PrintArea defined by a named range when building the print‑setup report.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintSetupReport
{
    // Creates a workbook, assigns sample PrintArea and PrintTitleRows values, adds a "PrintSetupReport" sheet, lists each worksheet’s name with its PrintArea and PrintTitleRows (or "Not Set"), and saves the workbook as PrintSetupReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample setup: create a few worksheets and define
            // print area / title rows for demonstration purposes
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";
            sheet1.PageSetup.PrintArea = "A1:D20";
            sheet1.PageSetup.PrintTitleRows = "$1:$2";

            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.PageSetup.PrintArea = "B2:G30";
            // No title rows set for this sheet

            Worksheet sheet3 = workbook.Worksheets.Add("Summary");
            // Neither print area nor title rows set for this sheet

            // -------------------------------------------------
            // Create a report worksheet that will list the settings
            // -------------------------------------------------
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "PrintSetupReport";

            // Header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Print Area");
            reportSheet.Cells[0, 2].PutValue("Print Title Rows");

            int reportRow = 1;

            // Iterate through all worksheets except the report sheet itself
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == reportSheet.Name)
                    continue;

                string printArea = ws.PageSetup.PrintArea;
                string titleRows = ws.PageSetup.PrintTitleRows;

                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                reportSheet.Cells[reportRow, 1].PutValue(string.IsNullOrEmpty(printArea) ? "Not Set" : printArea);
                reportSheet.Cells[reportRow, 2].PutValue(string.IsNullOrEmpty(titleRows) ? "Not Set" : titleRows);

                reportRow++;
            }

            // Save the workbook containing the report
            workbook.Save("PrintSetupReport.xlsx");

            Console.WriteLine("Report generated and saved as PrintSetupReport.xlsx");
        }
    }
}
