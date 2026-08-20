// Title: Hide Gridlines & Freeze Header Row in Excel using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds a header and sample rows, disables gridlines with IsGridlinesVisible = false, freezes the top row via FreezePanes, and saves the file as HideGridlinesAndFreezeHeader.xlsx.
// Keywords: Aspose.Cells hide gridlines C# | Aspose.Cells freeze header row | Aspose.Cells FreezePanes example | .NET Excel hide gridlines | Aspose.Cells workbook formatting
// Common Searches: hide gridlines Aspose.Cells .NET | freeze first row Aspose.Cells C# | Aspose.Cells hide gridlines and freeze panes | how to keep header visible while scrolling Excel Aspose
// Developer Intent: Remove worksheet gridlines and lock the header row in place for a cleaner, scroll‑friendly view.
// Use Cases: Produce a polished report without gridlines while keeping column titles fixed during scrolling. | Create a printable spreadsheet that hides gridlines but retains a static header for reference. | Design a data‑heavy dashboard where the header row stays visible as users navigate large datasets.
// AI Prompts: Show C# code to hide Excel gridlines and freeze the top row using Aspose.Cells. | Provide an Aspose.Cells .NET example that disables gridlines, freezes the header, and saves as .xlsx. | Explain the parameters of FreezePanes for locking the first row while keeping gridlines hidden.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a header and sample rows, disables gridlines with IsGridlinesVisible = false, freezes the top row via FreezePanes, and saves the file as HideGridlinesAndFreezeHeader.xlsx.
    public class HideGridlinesAndFreezeHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data
                worksheet.Cells["A1"].PutValue("Header");
                for (int i = 2; i <= 20; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                }

                // Hide gridlines
                worksheet.IsGridlinesVisible = false;

                // Freeze the first row (header)
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook
                workbook.Save("HideGridlinesAndFreezeHeader.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HideGridlinesAndFreezeHeader.Run();
        }
    }
}
