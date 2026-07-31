// Title: C# – Delete rows with null or DBNull in a required column using Aspose.Cells (reverse loop)
// Description: This Aspose.Cells example creates a workbook, adds sample data, then iterates from the last data row to the first (skipping the header) to detect null or DBNull values in a required column. Matching rows are removed with Cells.DeleteRow and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# delete rows | remove null rows Aspose.Cells | DBNull row deletion .NET | reverse loop delete Excel rows | Cells.DeleteRow example | Excel data cleansing C# | Aspose.Cells GitHub sample | worksheet row removal null column | C# Excel automation Aspose
// Common Searches: Aspose.Cells delete rows where column is null | C# remove rows with DBNull in Excel worksheet | How to loop backwards to delete rows in Aspose.Cells | Aspose.Cells example for cleaning data rows | Delete empty ID rows using Aspose.Cells .NET
// Developer Intent: Remove every worksheet row whose required column contains null or DBNull values.
// Use Cases: Clean imported datasets by discarding incomplete records before analysis. | Prepare Excel reports that require all rows to have a valid identifier. | Automate data validation in batch processing pipelines using Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that deletes rows where column A is null or DBNull, using a bottom‑up loop. | Explain how Cells.DeleteRow works when shifting rows after removal in Aspose.Cells. | Show how to adapt the loop to preserve cell formatting and formulas while deleting rows.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This Aspose.Cells example creates a workbook, adds sample data, then iterates from the last data row to the first (skipping the header) to detect null or DBNull values in a required column. Matching rows are removed with Cells.DeleteRow and the workbook is saved as an XLSX file.
    public class DeleteRowsWithNullInColumn
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: column A (index 0) is the required column
            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");

            // Data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");

            cells["A3"].PutValue(null);               // Null value in required column
            cells["B3"].PutValue("Bob");

            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Charlie");

            cells["A5"].PutValue(DBNull.Value);       // DBNull also considered null
            cells["B5"].PutValue("David");

            // Index of the required column (0‑based)
            int requiredColumnIndex = 0;

            // Loop from the last data row up to the first data row (skip header)
            for (int row = cells.MaxDataRow; row >= 1; row--)
            {
                Cell cell = cells[row, requiredColumnIndex];

                // Determine if the cell is null or DBNull
                bool isNull = cell.Value == null || cell.Value == DBNull.Value;

                if (isNull)
                {
                    // Delete the entire row and shift subsequent rows up
                    cells.DeleteRow(row, true);
                }
            }

            // Save the workbook
            string outputPath = "DeletedRows.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
