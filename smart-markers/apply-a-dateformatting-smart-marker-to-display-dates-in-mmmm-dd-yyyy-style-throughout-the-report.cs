// Title: Apply a 'MMMM dd, yyyy' date‑format smart marker to a column in an Excel report with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set custom date format in smart marker column | How to use WorkbookDesigner to format dates as 'MMMM dd, yyyy' in an Excel template | Applying date‑format smart marker to column B in Aspose.Cells report | Smart marker syntax for date formatting in Aspose.Cells .NET | Generate Excel report with formatted dates using Aspose.Cells and DataTable
// Tags: date format smart marker Aspose.Cells | WorkbookDesigner custom date formatting C# | Excel template column smart marker date styling | Aspose.Cells generate report with formatted dates | smart marker MMMM dd yyyy .NET

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel template, inserts the smart marker '{Date:MMMM dd, yyyy}' into cells B2‑B100, binds a DataTable of DateTime values, processes the markers with WorkbookDesigner, and saves the populated report as ReportOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook
                string templatePath = "ReportTemplate.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                var workbook = new Workbook(templatePath);
                var sheet = workbook.Worksheets[0];

                // Apply the date‑formatting smart marker to column B (index 1)
                int startRow = 1; // zero‑based index for row 2
                int endRow   = 99; // zero‑based index for row 100

                for (int row = startRow; row <= endRow; row++)
                {
                    sheet.Cells[row, 1].PutValue("{Date:MMMM dd, yyyy}");
                }

                // Prepare the data source
                DataTable dt = GetYourDataTable();

                // Process smart markers using WorkbookDesigner
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Save the resulting workbook
                string outputPath = "ReportOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Placeholder for actual data retrieval logic
        static DataTable GetYourDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("Date", typeof(DateTime));

            // Example rows (replace with real data)
            table.Rows.Add(DateTime.Now);
            table.Rows.Add(DateTime.Now.AddDays(1));
            table.Rows.Add(DateTime.Now.AddDays(2));

            return table;
        }
    }
}
