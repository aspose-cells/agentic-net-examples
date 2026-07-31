// Title: Aggregate Sales Data with Structured Reference Formulas in Aspose.Cells C#
// Description: Creates a workbook with two ListObject tables (SalesQ1 and SalesQ2) and a Summary sheet that uses structured reference formulas (e.g., =SUM(SalesQ1Table[Sales])) to calculate quarterly totals and a combined total, then recalculates and saves the file.
// Keywords: Aspose.Cells | C# | .NET | structured reference | ListObject | summary worksheet | aggregate tables | SUM formula | Excel automation | multiple tables aggregation
// Common Searches: Aspose.Cells sum column from ListObject | C# create summary sheet with structured references | aggregate data from multiple Excel tables using Aspose.Cells | how to use =SUM(Table[Column]) in Aspose.Cells | combine sales tables into one summary workbook C#
// Developer Intent: Create a workbook with two ListObject tables and a summary sheet that calculates each quarter’s total and the combined sales using structured reference formulas.
// Use Cases: Quarterly sales reporting with auto‑updating totals | Financial dashboard that consolidates data from several worksheets | Automated master summary for multi‑sheet Excel workbooks | Data validation and recalculation after adding new rows to source tables
// AI Prompts: Generate C# Aspose.Cells code that builds two worksheets with ListObject tables and adds a summary worksheet using structured reference formulas like =SUM(Table[Column]) | Show how to apply structured reference formulas, recalculate the workbook, and save the file in Aspose.Cells .NET | Provide an example of aggregating values from multiple tables into a single summary sheet with Aspose.Cells for C#

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject and ListObjectCollection

namespace AsposeCellsSummaryExample
{
    // Creates a workbook with two ListObject tables (SalesQ1 and SalesQ2) and a Summary sheet that uses structured reference formulas (e.g., =SUM(SalesQ1Table[Sales])) to calculate quarterly totals and a combined total, then recalculates and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ------------------------------
                // Worksheet 1 – First data table
                // ------------------------------
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "SalesQ1";

                // Populate sample data (header + 5 rows)
                ws1.Cells["A1"].PutValue("Region");
                ws1.Cells["B1"].PutValue("Sales");
                ws1.Cells["A2"].PutValue("North");
                ws1.Cells["B2"].PutValue(1200);
                ws1.Cells["A3"].PutValue("South");
                ws1.Cells["B3"].PutValue(950);
                ws1.Cells["A4"].PutValue("East");
                ws1.Cells["B4"].PutValue(780);
                ws1.Cells["A5"].PutValue("West");
                ws1.Cells["B5"].PutValue(660);
                ws1.Cells["A6"].PutValue("Central");
                ws1.Cells["B6"].PutValue(820);

                // Convert the range into a structured table (ListObject)
                ListObjectCollection tables1 = ws1.ListObjects;
                int tableIndex1 = tables1.Add(0, 0, 5, 1, true); // rows 0‑5, columns 0‑1
                ListObject table1 = tables1[tableIndex1];
                // Use DisplayName to set the table name (Name property not available in some versions)
                table1.DisplayName = "SalesQ1Table";

                // ------------------------------
                // Worksheet 2 – Second data table
                // ------------------------------
                Worksheet ws2 = workbook.Worksheets.Add("SalesQ2");

                // Populate sample data (header + 5 rows)
                ws2.Cells["A1"].PutValue("Region");
                ws2.Cells["B1"].PutValue("Sales");
                ws2.Cells["A2"].PutValue("North");
                ws2.Cells["B2"].PutValue(1100);
                ws2.Cells["A3"].PutValue("South");
                ws2.Cells["B3"].PutValue(1020);
                ws2.Cells["A4"].PutValue("East");
                ws2.Cells["B4"].PutValue(850);
                ws2.Cells["A5"].PutValue("West");
                ws2.Cells["B5"].PutValue(730);
                ws2.Cells["A6"].PutValue("Central");
                ws2.Cells["B6"].PutValue(910);

                // Convert the range into a structured table
                ListObjectCollection tables2 = ws2.ListObjects;
                int tableIndex2 = tables2.Add(0, 0, 5, 1, true);
                ListObject table2 = tables2[tableIndex2];
                table2.DisplayName = "SalesQ2Table";

                // ------------------------------
                // Summary worksheet – aggregation using structured references
                // ------------------------------
                Worksheet summaryWs = workbook.Worksheets.Add("Summary");

                // Header labels
                summaryWs.Cells["A1"].PutValue("Metric");
                summaryWs.Cells["B1"].PutValue("Value");

                // Total Sales Q1
                summaryWs.Cells["A2"].PutValue("Total Sales Q1");
                summaryWs.Cells["B2"].Formula = "=SUM(SalesQ1Table[Sales])";

                // Total Sales Q2
                summaryWs.Cells["A3"].PutValue("Total Sales Q2");
                summaryWs.Cells["B3"].Formula = "=SUM(SalesQ2Table[Sales])";

                // Combined Total Sales
                summaryWs.Cells["A4"].PutValue("Combined Total Sales");
                summaryWs.Cells["B4"].Formula = "=SUM(SalesQ1Table[Sales], SalesQ2Table[Sales])";

                // Recalculate all formulas so that the summary values are up‑to‑date
                workbook.CalculateFormula();

                // Save the workbook to a file
                string outputPath = "SummaryWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
