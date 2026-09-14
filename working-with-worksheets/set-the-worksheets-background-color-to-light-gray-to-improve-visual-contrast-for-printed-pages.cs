// Title: How to set an Excel worksheet tab color to light gray using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that changes the first worksheet's tab color to LightGray and saves the workbook. | Provide a snippet that demonstrates setting a worksheet's TabColor property to a custom shade for better print contrast in Aspose.Cells.
// Common Searches: Aspose.Cells C# change worksheet tab color to light gray for printing | set Excel sheet tab shade programmatically with Aspose.Cells .NET | how to improve printed worksheet contrast by adjusting tab color using Aspose.Cells
// Tags: Aspose.Cells worksheet TabColor customization | C# apply LightGray to worksheet tab | print-friendly worksheet tab shading Aspose.Cells | modify Excel tab appearance .NET | Aspose.Cells visual contrast for printed worksheets

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new Workbook (or loads an existing one), accesses the first Worksheet, sets its TabColor property to Color.LightGray to improve printed contrast, saves the file as 'WorksheetWithGrayBackground.xlsx', and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet tab color to light gray for better printed contrast
            worksheet.TabColor = Color.LightGray;

            // Save the workbook to a file
            string outputPath = "WorksheetWithGrayBackground.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
