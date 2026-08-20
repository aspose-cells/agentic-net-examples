// Title: Detect Empty Worksheets in Aspose.Cells (.NET) Using MaxDataRow & MaxDataColumn
// Description: C# example that creates a workbook, adds a populated sheet and an empty sheet, then checks each worksheet's Cells.MaxDataRow and Cells.MaxDataColumn. When both properties return -1 the sheet is identified as empty, the result is printed and the workbook saved.
// Keywords: Aspose.Cells empty worksheet detection | MaxDataRow -1 | MaxDataColumn -1 | C# check blank sheet Aspose.Cells | .NET workbook empty sheet | identify empty worksheet Aspose
// Common Searches: Aspose.Cells how to find empty worksheets | MaxDataRow and MaxDataColumn empty sheet .NET | C# detect blank worksheet in Aspose.Cells | list worksheets with no data Aspose.Cells
// Developer Intent: Determine which worksheets in a workbook contain no data by evaluating Cells.MaxDataRow and Cells.MaxDataColumn.
// Use Cases: Skip processing of blank sheets during report generation | Remove or hide empty worksheets to reduce file size | Validate workbook integrity by flagging sheets without values
// AI Prompts: Write a C# function that returns the names of all empty worksheets in an Aspose.Cells workbook using MaxDataRow and MaxDataColumn. | Provide code to delete every empty worksheet from a workbook after detection with Aspose.Cells. | Explain the behavior of MaxDataRow and MaxDataColumn when a sheet contains only formatting or comments but no cell values.

using System;
using Aspose.Cells;

namespace AsposeCellsEmptyWorksheetDetection
{
    // C# example that creates a workbook, adds a populated sheet and an empty sheet, then checks each worksheet's Cells.MaxDataRow and Cells.MaxDataColumn. When both properties return -1 the sheet is identified as empty, the result is printed and the workbook saved.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // For demonstration, a new empty workbook is created

            // Add a worksheet with data for testing
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataSheet";
            dataSheet.Cells["A1"].PutValue("Sample");

            // Add an empty worksheet
            Worksheet emptySheet = workbook.Worksheets.Add("EmptySheet");

            // Iterate through all worksheets and detect empty ones
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow and MaxDataColumn are -1 when the sheet contains no data
                bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                Console.WriteLine($"Worksheet \"{sheet.Name}\" is {(isEmpty ? "empty" : "not empty")}.");
            }

            // Optionally save the workbook to verify the result
            workbook.Save("EmptyWorksheetDetectionResult.xlsx");
        }
    }
}
