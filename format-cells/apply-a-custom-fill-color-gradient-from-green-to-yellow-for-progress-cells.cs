// Title: C# – Apply a Green‑to‑Yellow Vertical Gradient Fill to Excel Cells with Aspose.Cells
// Description: Demonstrates how to create a workbook, add task data, define a two‑color vertical gradient (green to yellow) using Aspose.Cells for .NET, apply it to the progress column, and save the file as ProgressGradient.xlsx.
// Keywords: Aspose.Cells C# gradient fill | two‑color gradient style .NET | vertical gradient Excel cells | cell background color scale | progress bar formatting Excel | Excel cell styling Aspose | gradient color formatting C#
// Common Searches: Aspose.Cells set vertical gradient background for cells | C# apply green to yellow gradient to Excel column | how to create progress bar effect with Aspose.Cells | two‑color gradient fill in Excel using .NET | apply gradient style to specific range Aspose.Cells
// Developer Intent: Create a vertical green‑to‑yellow two‑color gradient style and apply it to the progress cells in column B of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Show task completion percentages with a gradient‑filled progress column in project reports. | Generate a heat‑map style indicator for dashboards exported to Excel. | Add a colored progress‑bar visual for each row in status‑tracking spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells to apply a red‑to‑blue diagonal gradient to a selected range. | Show how to conditionally assign different gradient styles based on cell values in Aspose.Cells. | Provide an example of saving an Aspose.Cells workbook to a memory stream after applying gradient formatting.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add task data, define a two‑color vertical gradient (green to yellow) using Aspose.Cells for .NET, apply it to the progress column, and save the file as ProgressGradient.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data representing progress values
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["B3"].PutValue(50);
        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B4"].PutValue(80);

        // Create a style with a two‑color gradient from green to yellow
        Style gradientStyle = workbook.CreateStyle();
        gradientStyle.SetTwoColorGradient(Color.Green, Color.Yellow, GradientStyleType.Vertical, 1);

        // Apply the gradient style to the progress cells (column B)
        for (int row = 1; row <= 4; row++) // rows 2 to 5 (0‑based index)
        {
            Cell cell = sheet.Cells[row, 1]; // column index 1 corresponds to column B
            cell.SetStyle(gradientStyle);
        }

        // Save the workbook
        workbook.Save("ProgressGradient.xlsx");
    }
}
