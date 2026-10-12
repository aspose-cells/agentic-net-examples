// Title: Add a button control to an Excel worksheet and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, adds a Button shape at a specific row, column, and pixel size, sets its caption, and saves the file as XLSX with Aspose.Cells. | Show how to insert a button into an existing worksheet, optionally assign a macro or hyperlink, and persist the changes by saving the workbook. | Demonstrate positioning a button control using pixel coordinates on a worksheet and exporting the workbook while keeping the control intact.
// Common Searches: how to insert a button shape into an Excel file using Aspose.Cells C# | save workbook after adding form controls with Aspose.Cells .NET | set button text and assign macro in Aspose.Cells C# example | position Excel button in pixels using Aspose.Cells API | export worksheet with button control to XLSX using Aspose.Cells
// Tags: button shape insertion Aspose.Cells C# | persist workbook after shape addition Aspose.Cells | pixel based button placement Aspose.Cells | macro assignment to button Aspose.Cells | export worksheet with embedded button Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a button shape to the first worksheet with defined row, column, and pixel dimensions, optionally sets a macro or hyperlink, and saves the workbook as an XLSX file, preserving the new control.
class Program
{
    static void Main()
    {
        // Load an existing workbook (or create a new one)
        Workbook workbook = new Workbook(); // creates a new workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Add a button control to the worksheet
        // Define the position and size of the button (in pixels)
        int upperLeftRow = 2;
        int upperLeftColumn = 1;
        int top = 5;
        int left = 5;
        int height = 30;
        int width = 100;

        // Create the button
        Button button = sheet.Shapes.AddButton(upperLeftRow, upperLeftColumn, top, left, height, width);
        button.Text = "Click Me";

        // Optionally, assign a macro or hyperlink to the button
        // button.Macro = "MyMacro";
        // button.Hyperlink = new Hyperlink("http://example.com");

        // Save the workbook to preserve the added control
        workbook.Save("WorkbookWithButton.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved successfully with the new button control.");
    }
}
