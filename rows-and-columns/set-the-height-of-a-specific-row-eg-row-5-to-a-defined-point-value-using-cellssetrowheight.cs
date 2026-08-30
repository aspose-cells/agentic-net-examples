// Title: How to set the height of row 5 to 30 points in an Excel file using Aspose.Cells SetRowHeight (C#)
// AI Prompts: Write C# code that creates a workbook, sets row index 5 height to 30 points with Cells.SetRowHeight, and saves the file. | Show how to read back the height of a specific row after calling SetRowHeight in Aspose.Cells for .NET. | Explain the zero‑based row indexing rules when adjusting row height with the Aspose.Cells SetRowHeight method.
// Common Searches: Aspose.Cells C# set row height in points example | change height of row 5 in Excel using Aspose.Cells library | retrieve row height after setting it with Aspose.Cells .NET | how to use Cells.SetRowHeight for specific rows in a workbook | adjust Excel row height programmatically with Aspose.Cells C#
// Tags: Aspose.Cells SetRowHeight C# | Excel row height points Aspose.Cells | zero based row index Aspose.Cells | retrieve row height Aspose.Cells | save workbook after row height change Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightDemo
{
    // Demonstrates creating a workbook, accessing the first worksheet, setting row 5 (zero‑based) height to 30 points with Cells.SetRowHeight, retrieving the height, printing it, and saving the file as RowHeightDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection
            Cells cells = worksheet.Cells;

            // Set the height of row 5 (zero‑based index) to 30 points
            cells.SetRowHeight(5, 30);

            // Verify the height (optional)
            double height = cells.GetRowHeight(5);
            Console.WriteLine($"Row 5 height set to {height} points.");

            // Save the workbook
            workbook.Save("RowHeightDemo.xlsx");
        }
    }
}
