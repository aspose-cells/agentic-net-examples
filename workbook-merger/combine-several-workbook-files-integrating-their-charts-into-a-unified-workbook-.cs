using System;
using Aspose.Cells;

namespace WorkbookCombineDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Report_Q1.xlsx",
                "Report_Q2.xlsx",
                "Report_Q3.xlsx"
            };

            // Destination workbook – created with the default constructor
            Workbook combinedWorkbook = new Workbook();

            // Load each source workbook and combine it into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load source workbook using the string constructor (load rule)
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook (combine rule)
                combinedWorkbook.Combine(sourceWorkbook);
            }

            // Refresh all pivot tables and charts to ensure relationships are up‑to‑date
            combinedWorkbook.Worksheets.RefreshAll();

            // Save the unified workbook (save rule)
            combinedWorkbook.Save("CombinedReport.xlsx", SaveFormat.Xlsx);
        }
    }
}