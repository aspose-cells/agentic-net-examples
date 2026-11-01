// Title: List all cells monitored by the Watch Window in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# code to retrieve watched cells from Aspose.Cells WatchWindow | How to enumerate WatchWindowCell objects in Aspose.Cells .NET | Get addresses of cells monitored by Watch Window with Aspose.Cells | Aspose.Cells dynamic worksheet WatchWindow example C# | Extract WatchWindow cell list from workbook using Aspose.Cells
// Tags: Aspose.Cells WatchWindow cell enumeration C# | retrieve WatchWindowCell addresses Aspose.Cells | dynamic worksheet access Aspose.Cells .NET | zero based index to A1 address conversion C# | list monitored cells Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

// The example loads an Excel workbook, accesses the first worksheet as a dynamic object to obtain its WatchWindow, iterates through the WatchWindowCells collection, converts each cell's zero‑based row and column indices to an A1‑style address, and outputs the worksheet name and cell address.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet as a dynamic object to allow late‑bound members (e.g., WatchWindow)
            dynamic worksheet = workbook.Worksheets[0];
            dynamic watchWindow;

            try
            {
                // Attempt to retrieve the WatchWindow; may throw if the feature is unavailable in this version
                watchWindow = worksheet.WatchWindow;
            }
            catch (RuntimeBinderException)
            {
                Console.WriteLine("WatchWindow feature is not supported in the current Aspose.Cells version.");
                return;
            }

            // Iterate through all cells monitored by the Watch Window
            foreach (dynamic watchedCell in watchWindow.WatchWindowCells)
            {
                // Worksheet index of the watched cell
                int wsIndex = watchedCell.WorksheetIndex;
                string sheetName = workbook.Worksheets[wsIndex].Name;

                // Convert zero‑based row/column to A1 address
                string address = CellReference(watchedCell.Row, watchedCell.Column);
                Console.WriteLine($"Worksheet: {sheetName}, Cell: {address}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper: convert zero‑based row/column indices to A1 style address
    static string CellReference(int row, int column)
    {
        int dividend = column + 1;
        string columnName = "";
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{columnName}{row + 1}";
    }
}
