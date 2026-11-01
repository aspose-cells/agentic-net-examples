// Title: Add a rectangle shape with a hyperlink to a Summary sheet cell on every worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, creates a "Summary" worksheet if it does not exist, then iterates over all other worksheets to place a rectangle shape at row 5 column 5, set its displayed text, and attach a hyperlink to cell A1 on the Summary sheet. | Refactor the shape‑creation logic so the drawing type (e.g., Rectangle, Oval, RoundedRectangle) and the target cell on the Summary sheet are supplied via parameters, while preserving per‑worksheet try/catch blocks and console logging. | Extend the program to write a detailed log file that records, for each worksheet, whether the shape was added successfully or an exception was thrown, and then save the modified workbook to a user‑specified output path.
// Common Searches: how to add a shape to each worksheet in an Excel file using Aspose.Cells C# | Aspose.Cells C# loop worksheets and create hyperlink from shape to another sheet | C# Aspose.Cells add a rectangular drawing linked to a summary cell | skip a specific worksheet when adding shapes with Aspose.Cells .NET | log shape insertion results per worksheet Aspose.Cells example
// Tags: Aspose.Cells rectangle shape insertion across worksheets | shape hyperlink to summary sheet cell Aspose.Cells | iterate worksheets excluding specific sheet C# | Aspose.Cells error handling during shape addition | C# workbook save after modifying shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads or creates a workbook, ensures a "Summary" worksheet exists, then loops through every other worksheet, adds a rectangle shape at row 5/column 5, sets its text, links it to cell A1 on the Summary sheet, logs success or failure for each sheet, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook has been created.");
            }

            // Ensure a summary worksheet exists (create if missing)
            Worksheet summarySheet = workbook.Worksheets["Summary"];
            if (summarySheet == null)
            {
                int summaryIndex = workbook.Worksheets.Add();
                summarySheet = workbook.Worksheets[summaryIndex];
                summarySheet.Name = "Summary";
                summarySheet.Cells["A1"].PutValue("Summary Information");
            }

            // Iterate through all worksheets except the summary sheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                    continue; // Skip the summary sheet itself

                try
                {
                    // Add a rectangle shape to the current worksheet
                    // Parameters: drawing type, upper-left row, upper-left column, top, left, height, width
                    Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 100, 50);

                    // Set the shape's displayed text
                    shape.Text = "Link to Summary";

                    // Attach a hyperlink that points to cell A1 on the Summary sheet
                    shape.AddHyperlink($"{summarySheet.Name}!A1");

                    Console.WriteLine($"Added shape to worksheet '{sheet.Name}' linking to '{summarySheet.Name}!A1'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add shape to worksheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
