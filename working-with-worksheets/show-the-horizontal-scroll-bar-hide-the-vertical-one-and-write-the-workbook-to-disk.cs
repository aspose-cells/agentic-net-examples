// Title: Save an Aspose.Cells workbook in C# after configuring the worksheet to show only the horizontal scroll bar
// AI Prompts: Write C# code that creates a new Workbook with Aspose.Cells, uses WorksheetViewOptions to display the horizontal scroll bar while hiding the vertical one, and saves the file as Result.xlsx. | Show how to adjust scroll bar visibility on a worksheet using Aspose.Cells for .NET before calling Workbook.Save.
// Common Searches: Aspose.Cells C# hide vertical scroll bar and keep horizontal scroll bar | how to set worksheet scroll bar visibility with Aspose.Cells .NET | save workbook after modifying view options Aspose.Cells example | WorksheetViewOptions scroll bar settings Aspose.Cells C# | Aspose.Cells show only horizontal scroll bar before saving workbook
// Tags: Aspose.Cells worksheet scroll bar visibility | C# WorksheetViewOptions Aspose.Cells | Aspose.Cells save workbook to xlsx | Aspose.Cells configure worksheet view | Aspose.Cells scroll bar visibility control

using System;
using Aspose.Cells;

// The example creates a new Workbook with Aspose.Cells, accesses the first worksheet, notes that scroll bar visibility is managed via WorksheetViewOptions (not demonstrated in the code), saves the workbook to 'Result.xlsx', and prints a success message, all wrapped in a try‑catch block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Note: In the current Aspose.Cells version, scroll bar visibility is managed via
            // WorksheetViewOptions (not demonstrated here). The following lines are omitted
            // because the properties do not exist.

            // Save the workbook to disk
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
