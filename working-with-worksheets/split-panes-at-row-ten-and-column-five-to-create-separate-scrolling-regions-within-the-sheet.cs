// Title: Split worksheet panes at row 10 & column 5 with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, populating sample data, splitting the worksheet window, setting the bottom pane to start at row 11 and the right pane at column 6, and saving the file as an XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells split panes | C# split worksheet | set split row column Aspose.Cells | PaneCollection FirstVisibleRowOfBottomPane | FirstVisibleColumnOfRightPane | Excel split panes .NET | freeze panes programmatically
// Common Searches: Aspose.Cells split panes C# | How to set split row and column in Aspose.Cells | Freeze panes at row 10 column 5 using Aspose.Cells | Programmatically split Excel window with Aspose.Cells .NET | Set first visible row of bottom pane Aspose.Cells
// Developer Intent: Programmatically split a worksheet into separate scrolling panes at row 10 and column 5 using Aspose.Cells for .NET.
// Use Cases: Keep header rows and columns visible while the rest of the sheet scrolls independently. | Create a static top‑left quadrant for titles and allow independent scrolling of the data area. | Design dashboards where navigation panels remain fixed and data sections scroll separately.
// AI Prompts: Write C# code with Aspose.Cells to split panes at a specific row and column and save the workbook. | Explain how to adjust split pane positions dynamically based on worksheet dimensions using Aspose.Cells. | Show how to read, modify, or remove existing pane settings in an existing Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, populating sample data, splitting the worksheet window, setting the bottom pane to start at row 11 and the right pane at column 6, and saving the file as an XLSX using Aspose.Cells for .NET.
public class SplitPaneDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data to visualize the split
            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j < 15; j++)
                {
                    worksheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Split the worksheet window
            worksheet.Split();

            // Set split positions
            PaneCollection panes = worksheet.GetPanes();
            panes.FirstVisibleRowOfBottomPane = 10;   // Split after row 10 (zero‑based index)
            panes.FirstVisibleColumnOfRightPane = 5; // Split after column 5 (zero‑based index)

            // Save the workbook
            workbook.Save("SplitPaneDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        SplitPaneDemo.Run();
    }
}
