// Title: Generate an Excel documentation workbook listing all global and worksheet‑scoped named ranges with their RefersTo formulas using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells list all named ranges and their RefersTo formulas | how to export global and local named ranges to a new Excel file using Aspose.Cells | create a named range report workbook in .NET with Aspose.Cells | retrieve worksheet‑scoped named range definitions programmatically in C# | Aspose.Cells generate documentation of named ranges in Excel
// Tags: Aspose.Cells enumerate named ranges in .NET | export named range RefersTo to Excel workbook | global and local named range extraction C# | named range documentation generation Aspose.Cells | dynamic Names collection handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeReport
{
    // The example loads a source workbook, creates a new workbook for reporting, writes header cells, iterates through all global named ranges via the workbook's Names collection, then iterates each worksheet's local named ranges, records each name (prefixed with the sheet name for locals) and its RefersTo formula into the report sheet, auto‑fits columns, and saves the documentation workbook while handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths (adjust as needed)
            string sourceFilePath = @"C:\Path\To\SourceWorkbook.xlsx";
            string reportFilePath = @"C:\Path\To\NamedRangeReport.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file not found: {sourceFilePath}");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourceFilePath);

                // Create a new workbook for the report
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];

                // Write header titles
                reportSheet.Cells["A1"].PutValue("Named Range");
                reportSheet.Cells["B1"].PutValue("Refers To");

                int rowIndex = 1; // Start after header (0‑based)

                // ---- Global named ranges ----
                // Use dynamic to access Names collection in case of version differences
                dynamic wbDynamic = sourceWorkbook;
                foreach (var globalName in wbDynamic.Names)
                {
                    // Access properties via dynamic as well
                    string name = globalName.Name;
                    string refersTo = globalName.RefersTo;

                    reportSheet.Cells[rowIndex, 0].PutValue(name);
                    reportSheet.Cells[rowIndex, 1].PutValue(refersTo);
                    rowIndex++;
                }

                // ---- Local (worksheet‑scoped) named ranges ----
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    dynamic wsDynamic = ws;
                    foreach (var localName in wsDynamic.Names)
                    {
                        string fullName = $"{ws.Name}!{localName.Name}";
                        string refersTo = localName.RefersTo;

                        reportSheet.Cells[rowIndex, 0].PutValue(fullName);
                        reportSheet.Cells[rowIndex, 1].PutValue(refersTo);
                        rowIndex++;
                    }
                }

                // Auto‑fit columns for readability
                reportSheet.AutoFitColumns();

                // Save the report workbook
                reportWorkbook.Save(reportFilePath);
                Console.WriteLine($"Report saved to: {reportFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
