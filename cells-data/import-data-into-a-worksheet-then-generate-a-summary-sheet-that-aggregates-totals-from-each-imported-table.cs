// Title: Import multiple DataTables into separate worksheets and generate a summary sheet that aggregates total Amount values with Aspose.Cells for .NET
// AI Prompts: Import two DataTable objects into distinct worksheets, display column headers, and add a summary worksheet that uses SUM formulas to calculate each sheet's total Amount. | Create a workbook, auto‑fit all columns after importing the tables, and save the file as an XLSX document using Aspose.Cells. | Configure ImportTableOptions to show field names and insert rows when importing DataTables into worksheets.
// Common Searches: how to import a DataTable into a specific worksheet with headers using Aspose.Cells C# | asp.net create summary worksheet that sums a column from multiple sheets in Aspose.Cells | set SUM formula for a column across worksheets with Aspose.Cells API
// Tags: ImportTableOptions show field names Aspose.Cells | DataTable to worksheet import Aspose.Cells C# | summary worksheet SUM formula across sheets Aspose.Cells | auto fit columns after data import Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportAndSummarize
{
    // The program creates a new workbook, imports two DataTable objects into separate worksheets with column headers, adds a summary worksheet that uses SUM formulas to total the Amount column from each sheet, auto‑fits the columns for readability, and saves the workbook as ImportedAndSummarized.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Import first table ----------
            // Prepare sample DataTable
            DataTable table1 = new DataTable("Sales_Q1");
            table1.Columns.Add("Product", typeof(string));
            table1.Columns.Add("Quantity", typeof(int));
            table1.Columns.Add("Amount", typeof(double));

            // Add rows
            table1.Rows.Add("Apple", 120, 1500.0);
            table1.Rows.Add("Banana", 80, 800.0);
            table1.Rows.Add("Cherry", 50, 1250.0);

            // Set import options (show column headers)
            ImportTableOptions options1 = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };

            // Import into the first worksheet starting at cell A1 (row 0, column 0)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Q1";
            sheet1.Cells.ImportData(table1, 0, 0, options1);

            // ---------- Import second table ----------
            // Create a new worksheet for the second table
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Q2";

            DataTable table2 = new DataTable("Sales_Q2");
            table2.Columns.Add("Product", typeof(string));
            table2.Columns.Add("Quantity", typeof(int));
            table2.Columns.Add("Amount", typeof(double));

            table2.Rows.Add("Apple", 100, 1300.0);
            table2.Rows.Add("Banana", 90, 950.0);
            table2.Rows.Add("Cherry", 70, 1400.0);

            ImportTableOptions options2 = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };

            // Import into the second worksheet starting at cell A1
            sheet2.Cells.ImportData(table2, 0, 0, options2);

            // ---------- Create summary sheet ----------
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Header row
            summarySheet.Cells["A1"].PutValue("Source Sheet");
            summarySheet.Cells["B1"].PutValue("Total Amount");

            // Row for first table total
            summarySheet.Cells["A2"].PutValue(sheet1.Name);
            // Formula to sum the "Amount" column of Q1 (excluding header)
            // Assuming "Amount" is column C (index 2) and data starts at row 2 (index 1)
            summarySheet.Cells["B2"].Formula = $"SUM('{sheet1.Name}'!C2:C{table1.Rows.Count + 1})";

            // Row for second table total
            summarySheet.Cells["A3"].PutValue(sheet2.Name);
            summarySheet.Cells["B3"].Formula = $"SUM('{sheet2.Name}'!C2:C{table2.Rows.Count + 1})";

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("ImportedAndSummarized.xlsx", SaveFormat.Xlsx);
        }
    }
}
