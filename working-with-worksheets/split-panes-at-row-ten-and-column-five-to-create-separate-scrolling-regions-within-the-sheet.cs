// Title: C# – Split Excel worksheet panes at row 10 and column 5 using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate sample data, enable pane splitting with Worksheet.Split(), and position the bottom and right panes at row 10 and column 5 via PaneCollection. The file is saved as SplitPaneDemo.xlsx.
// Keywords: Aspose.Cells split panes C# | Worksheet.Split Aspose.Cells | PaneCollection FirstVisibleRowOfBottomPane | FirstVisibleColumnOfRightPane | Excel split view programmatically | freeze panes vs split panes Aspose | Aspose.Cells .NET API pane splitting
// Common Searches: how to split Excel panes with Aspose.Cells C# | set bottom pane start row Aspose.Cells | split worksheet at specific column using Aspose | Aspose.Cells example for pane splitting | C# code to create scrollable panes in Excel
// Developer Intent: Programmatically divide an Excel sheet into independent scrolling areas by splitting after row 10 and column 5.
// Use Cases: Show a fixed header (rows 1‑10) while allowing vertical scrolling of the data below. | Separate a summary block (columns A‑E) from detailed columns, enabling independent horizontal navigation. | Create a large report where both top rows and left columns stay visible in distinct panes for quick reference.
// AI Prompts: Generate C# code to split worksheet panes at a custom row and column with Aspose.Cells. | Explain the differences between Worksheet.Split() and Worksheet.FreezePanes() in Aspose.Cells. | Show how to apply a different zoom level to each pane after splitting using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, populate sample data, enable pane splitting with Worksheet.Split(), and position the bottom and right panes at row 10 and column 5 via PaneCollection. The file is saved as SplitPaneDemo.xlsx.
public class SplitPaneDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some data to visualize the split
            for (int i = 0; i < 30; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Split the worksheet window
            worksheet.Split();

            // Set split positions
            PaneCollection panes = worksheet.GetPanes();
            panes.FirstVisibleRowOfBottomPane = 10;   // Horizontal split after row 10
            panes.FirstVisibleColumnOfRightPane = 5; // Vertical split after column 5

            // Save the workbook
            string outputPath = "SplitPaneDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
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
