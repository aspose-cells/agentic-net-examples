// Title: C# – Apply dd‑MMM‑yyyy Custom Date Format to an Aspose.Cells Table Column
// Description: Shows how to create a workbook, import a DataTable, turn the range into a ListObject, define a style with the custom format dd‑MMM‑yyyy, and apply that style to the OrderDate column of the Excel table using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | custom date format | dd-MMM-yyyy | Excel table | ListObject | style | SetStyle | ImportData | DataTable | TableColumnCustomDateFormat | column formatting | date styling | Excel export
// Common Searches: Aspose.Cells set custom date format for table column | dd-MMM-yyyy format ListObject C# | apply date style to Excel table column Aspose.Cells | how to format dates in an imported DataTable using Aspose.Cells | C# Aspose.Cells table column date formatting
// Developer Intent: Display the OrderDate column of an Aspose.Cells ListObject using the dd‑MMM‑yyyy pattern.
// Use Cases: Standardize date appearance in financial or sales reports generated from DataTables. | Ensure consistent date formatting for downstream data processing or BI tools. | Create a reusable date style that can be applied to multiple tables within the same workbook.
// AI Prompts: Generate C# code that sets a custom date format for a specific ListObject column without iterating over each cell. | Explain how to create and reuse a named style for date columns across several Aspose.Cells tables. | Show how to apply the dd‑MMM‑yyyy format to a table column when importing data from a DataTable in Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, import a DataTable, turn the range into a ListObject, define a style with the custom format dd‑MMM‑yyyy, and apply that style to the OrderDate column of the Excel table using Aspose.Cells for .NET.
    public class TableColumnCustomDateFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Prepare sample data with a DateTime column
                DataTable dt = new DataTable();
                dt.Columns.Add("OrderDate", typeof(DateTime));
                dt.Columns.Add("Product", typeof(string));
                dt.Rows.Add(new DateTime(2023, 1, 15), "Apple");
                dt.Rows.Add(new DateTime(2023, 2, 20), "Banana");
                dt.Rows.Add(new DateTime(2023, 3, 25), "Cherry");

                // Import the DataTable into the worksheet starting at A1 (including header)
                cells.ImportData(dt, 0, 0, new ImportTableOptions()
                {
                    IsFieldNameShown = true
                });

                // Convert the imported range into a ListObject (Excel table)
                // The data occupies rows 0‑3 and columns 0‑1 (A1:B4)
                int tableIndex = worksheet.ListObjects.Add(0, 0, dt.Rows.Count, dt.Columns.Count - 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "Orders";

                // Create a style with the desired custom date format
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";

                // Apply the style to the entire data column of the table (OrderDate column)
                int firstDataRow = table.DataRange.FirstRow + 1; // first row after header
                int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
                int dateColumnIndex = table.DataRange.FirstColumn; // OrderDate is the first column

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    cells[row, dateColumnIndex].SetStyle(dateStyle);
                }

                // Save the workbook
                workbook.Save("TableColumnCustomDateFormat.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TableColumnCustomDateFormat.Run();
        }
    }
}
