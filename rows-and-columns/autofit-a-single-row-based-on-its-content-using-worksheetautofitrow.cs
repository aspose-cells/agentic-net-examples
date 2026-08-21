// Title: C# – Auto‑fit a single worksheet row with wrapped text using Aspose.Cells
// Description: This .NET example creates a workbook, writes long strings to cells A1 and B1, enables text wrapping, then calls Worksheet.AutoFitRow(0) to automatically adjust the row height before saving the file.
// Keywords: Aspose.Cells | Worksheet.AutoFitRow | C# Excel automation | row height auto fit | wrap text in Excel | Aspose.Cells sample | Excel row auto‑adjust | GitHub example | global
// Common Searches: How to auto‑fit a row height in Aspose.Cells C# | Worksheet.AutoFitRow with text wrapping | Increase Excel row height programmatically .NET | Aspose.Cells row auto‑fit example | C# adjust row height to fit wrapped text
// Developer Intent: Resize a specific row so its wrapped content is fully visible.
// Use Cases: Generating reports where header rows contain lengthy descriptions | Creating invoices where product notes span multiple lines | Exporting data with variable‑length comments that need dynamic row heights | Building dashboards that automatically adapt row size to content
// AI Prompts: Provide C# code to auto‑fit multiple rows after setting wrap text with Aspose.Cells. | Show how to auto‑fit both columns and rows together in a .NET workbook. | Explain how to exclude a particular row from auto‑fit while applying it to others using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This .NET example creates a workbook, writes long strings to cells A1 and B1, enables text wrapping, then calls Worksheet.AutoFitRow(0) to automatically adjust the row height before saving the file.
public class AutoFitSingleRowDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add long text to cells in the first row to demonstrate autofit
            worksheet.Cells["A1"].PutValue("This is a very long text that should cause the row height to increase when auto-fitted.");
            worksheet.Cells["B1"].PutValue("Another long piece of text in the same row.");

            // Enable text wrapping so the row height can expand
            Style styleA = worksheet.Cells["A1"].GetStyle();
            styleA.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(styleA);

            Style styleB = worksheet.Cells["B1"].GetStyle();
            styleB.IsTextWrapped = true;
            worksheet.Cells["B1"].SetStyle(styleB);

            // Auto‑fit the first row (row index 0)
            worksheet.AutoFitRow(0);

            // Ensure output directory exists
            string outputPath = "AutoFitSingleRowDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during AutoFitSingleRowDemo: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            AutoFitSingleRowDemo.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
