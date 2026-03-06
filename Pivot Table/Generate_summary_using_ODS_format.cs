using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsSummaryOdsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that will be summarized
            string sourcePath = "input.xlsx";

            // -----------------------------------------------------------------
            // 1. Generate a simple textual summary of the spreadsheet
            // -----------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook(sourcePath);
            int sheetCount = sourceWorkbook.Worksheets.Count;
            int totalRows = 0;
            int totalColumns = 0;

            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                totalRows += ws.Cells.MaxDataRow + 1;          // MaxDataRow is zero‑based
                totalColumns += ws.Cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based
            }

            string summary = $"The workbook contains {sheetCount} worksheet(s) with a total of {totalRows} row(s) and {totalColumns} column(s).";

            // -----------------------------------------------------------------
            // 2. Create a new workbook to hold the summary text
            // -----------------------------------------------------------------
            Workbook summaryWorkbook = new Workbook();               // create workbook
            Worksheet sheet = summaryWorkbook.Worksheets[0];        // get first worksheet

            // Write the summary into cell A1
            sheet.Cells["A1"].PutValue(summary);

            // -----------------------------------------------------------------
            // 3. Configure ODS save options
            // -----------------------------------------------------------------
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
            // OdfVersion setting is optional; omitted for compatibility

            // -----------------------------------------------------------------
            // 4. Save the workbook as an ODS file using the configured options
            // -----------------------------------------------------------------
            string outputPath = "summary_output.ods";
            summaryWorkbook.Save(outputPath, saveOptions);          // save workbook

            Console.WriteLine($"Spreadsheet summary saved to ODS file: {outputPath}");
        }
    }
}