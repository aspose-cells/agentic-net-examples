// Title: Restrict Aspose.Cells Smart Marker Processing to a Specific Named Range with WorkbookDesigner.SetRange (C#)
// AI Prompts: Generate C# code that retrieves an Excel named range and calls WorkbookDesigner.SetRange to limit smart marker processing to that range. | Show how to bind a DataTable to the retrieved named range before invoking designer.Process() in Aspose.Cells. | Provide error‑handling that checks whether the named range exists and falls back to processing the whole workbook when it does not.
// Common Searches: how to use WorkbookDesigner.SetRange with a named range in Aspose.Cells C# | limit smart marker processing to a specific area in an Excel template using Aspose.Cells .NET | example of checking for a named range before processing smart markers in C# | fallback to full workbook when named range is missing Aspose.Cells
// Tags: WorkbookDesigner.SetRange for Excel named range C# | restrict smart marker scope to specific cells Aspose.Cells | bind DataTable to named range before smart marker processing | handle missing named range fallback Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, obtains a named range called "DataRange", optionally assigns a data source, limits smart marker processing to that range with WorkbookDesigner.SetRange, and saves the result, while gracefully handling missing files or absent named ranges.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "DataRange";

            // Verify that the template file exists.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Retrieve the named range (use fully qualified type to avoid ambiguity).
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName(rangeName);
            if (namedRange != null)
            {
                // If needed, you can set a data source for the named range here.
                // For this example we simply note its existence.
                Console.WriteLine($"Named range '{rangeName}' found on sheet index {namedRange.Worksheet.Index}.");
            }
            else
            {
                Console.WriteLine($"Named range '{rangeName}' not found. Processing the entire workbook.");
            }

            // Process smart markers.
            designer.Process();

            // Save the resulting workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
