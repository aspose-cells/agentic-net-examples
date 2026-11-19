// Title: Add a Form Control button to an Excel worksheet and save the workbook using Aspose.Cells for .NET
// Common Searches: how to insert a form control button into an Excel sheet with Aspose.Cells C# | saving changes to shapes after adding a button in Aspose.Cells .NET | Aspose.Cells C# add button control and preserve it on workbook save | create or load a workbook then add a form control button using Aspose.Cells | Aspose.Cells example for adding a button shape and saving the file
// Tags: form control button insertion Aspose.Cells | worksheet shape addition C# | save workbook with updated shapes Aspose.Cells | retain Excel controls after save Aspose.Cells | button size configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file (or creates a new workbook), adds a Form Control button shape to the first worksheet with specified position and dimensions, sets its caption, and then saves the workbook to preserve the newly added control.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add a new button control (Form Control) to the sheet
            // -------------------------------------------------
            // Parameters: upper left row, upper left column,
            // upper left row offset (pixels), upper left column offset (pixels),
            // height (pixels), width (pixels)
            Button button = sheet.Shapes.AddButton(
                1,          // upper left row (zero‑based)
                1,          // upper left column (zero‑based)
                0,          // upper left row offset
                0,          // upper left column offset
                50,         // height
                100);       // width

            // Set the button's display text.
            button.Text = "Click Me";

            // -------------------------------------------------
            // Save the workbook to preserve the added/updated controls
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
