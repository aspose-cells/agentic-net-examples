// Title: C# example: Loop through RichTextPortion objects in a worksheet cell and set each portion's font name using Aspose.Cells
// AI Prompts: Generate C# code that uses Aspose.Cells to enumerate the RichTextPortion collection of cell A1 and assign a new Font.Name to every portion while preserving existing formatting. | Provide a step‑by‑step snippet that loads an Excel file, accesses a cell’s RichTextPortion objects, changes their font to "Times New Roman", and saves the workbook.
// Common Searches: aspnet cells iterate richtextportion collection change font name c# | how to modify font of each text segment in an Excel cell using Aspose.Cells | c# aspose.cells set different fonts for parts of a cell's text | example code for updating RichTextPortion font in worksheet cell Aspose.Cells | loop through cell rich text portions and apply new font Aspose.Cells .NET
// Tags: richtextportion font update Aspose.Cells | enumerate cell text portions .NET | set font name for each rich text segment Excel | Aspose.Cells modify individual text portion styling | C# enumerate RichTextPortion collection | Excel cell rich text formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook (or creates a new one), ensures cell A1 contains text, accesses the cell’s RichTextPortion collection, iterates through each portion to assign a new Font.Name, and saves the modified workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet and target cell (A1).
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Ensure the cell has some text; if empty, set a default value.
            if (string.IsNullOrEmpty(cell.StringValue))
            {
                cell.PutValue("Sample Text");
            }

            // Apply Calibri font to the entire cell text.
            Style style = cell.GetStyle();
            style.Font.Name = "Calibri";
            cell.SetStyle(style);

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
