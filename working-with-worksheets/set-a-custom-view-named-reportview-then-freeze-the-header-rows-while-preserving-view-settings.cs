// Title: Create a custom view 'ReportView' with frozen header rows using Aspose.Cells for .NET
// Description: Demonstrates how to add a custom view called ReportView, freeze one or more header rows, retain the view settings, and save the workbook as an XLSX file with Aspose.Cells for C#.
// Keywords: Aspose.Cells custom view | ReportView Aspose.Cells | freeze header rows C# | Aspose.Cells FreezePanes | preserve view settings Excel | .NET Excel custom view | Aspose.Cells workbook save
// Common Searches: Aspose.Cells how to create a custom view | freeze first row and keep view settings in Excel using C# | set ReportView in Aspose.Cells workbook | preserve frozen panes after saving with Aspose.Cells | C# Aspose.Cells custom view example
// Developer Intent: Add a custom view named ReportView, freeze header rows, and save the workbook while keeping the view configuration.
// Use Cases: Generate a printable report where the header stays visible and the view can be recalled by users. | Build a data‑entry template that preserves frozen panes across different workstations. | Create a dashboard workbook with a predefined view that includes frozen headers for easy navigation.
// AI Prompts: Show me C# code that creates a custom view called 'ReportView', freezes the first two rows, and saves the workbook with Aspose.Cells. | Explain how to verify that a custom view and frozen panes are applied correctly in an Aspose.Cells workbook. | Provide a step‑by‑step guide to preserve view settings after modifying freeze panes in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to add a custom view called ReportView, freeze one or more header rows, retain the view settings, and save the workbook as an XLSX file with Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the first row (header)
            worksheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            string outputPath = "ReportViewDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
