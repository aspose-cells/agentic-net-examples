// Title: Create a named range for the visible rows of a filtered ListObject and use it in a SUBTOTAL formula with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a ListObject, applies an AutoFilter, defines a workbook‑level named range that points to the visible Amount cells, and inserts a SUBTOTAL(9, ...) formula referencing that range. | Show how to generate a SUBTOTAL calculation for only the displayed rows of a table by creating a structured named range after filtering, then assign the formula to a cell in a .NET Excel workbook.
// Common Searches: aspnet create named range for filtered table column using Aspose.Cells | c# subtotal formula for visible rows after applying AutoFilter with Aspose.Cells | aspose.cells structured reference for filtered ListObject in SUBTOTAL | generate Excel workbook in C# that sums only filtered data using a named range
// Tags: Aspose.Cells define named range from filtered ListObject | C# SUBTOTAL formula with structured reference | Aspose.Cells AutoFilter visible rows subtotal | Excel workbook named range for filtered column .NET | Aspose.Cells ListObject subtotal calculation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds sample Category and Amount data, converts the range into a ListObject named "SalesData", filters the table to show only rows where Category equals "A", defines a workbook‑level named range that includes the visible Amount cells, inserts a SUBTOTAL(9, ...) formula that references this named range, and saves the file as FilteredTableWithSubtotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data (a simple table with Category and Amount)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");

            string[,] data = new string[,]
            {
                { "A", "100" },
                { "B", "200" },
                { "A", "150" },
                { "C", "300" },
                { "B", "250" }
            };

            int rows = data.GetLength(0);
            for (int i = 0; i < rows; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(data[i, 0]); // Category column (A)
                sheet.Cells[i + 1, 1].PutValue(Convert.ToDouble(data[i, 1])); // Amount column (B)
            }

            // Convert the range into a table (ListObject)
            int totalRows = rows + 1; // include header row
            int totalCols = 2;        // Category and Amount
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows, totalCols, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesData"; // Set table name
            table.ShowHeaderRow = true;
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Apply an AutoFilter to show only rows where Category = "A"
            table.AutoFilter.Filter(0, "A"); // 0 = first column of the table (Category)

            // Insert a SUBTOTAL formula that sums the visible (filtered) amounts
            sheet.Cells["D1"].PutValue("Subtotal of Visible Amounts:");
            sheet.Cells["E1"].Formula = "=SUBTOTAL(9, SalesData[Amount])";

            // Ensure output directory exists
            string outputPath = "FilteredTableWithSubtotal.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
