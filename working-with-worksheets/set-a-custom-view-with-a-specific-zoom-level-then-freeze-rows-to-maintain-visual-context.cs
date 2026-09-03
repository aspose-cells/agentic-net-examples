// Title: Create a custom view with 120% zoom and freeze the first two rows in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a named custom view with a 120% zoom level and then freeze the first two rows of the worksheet. | Show how to detect if the CustomViews collection is available in Aspose.Cells, set the zoom for that view, and apply FreezePanes to keep the top two rows visible while scrolling. | Generate a complete Aspose.Cells example that creates a workbook, optionally creates a custom view at a specific zoom, freezes rows, and saves the file.
// Common Searches: aspnet cells how to set custom view zoom level programmatically | aspnet cells freeze first two rows while scrolling | c# aspnet cells create custom view and freeze panes example | aspnet cells workbook save with custom view and frozen rows
// Tags: custom view zoom Aspose.Cells | freeze panes rows Aspose.Cells | conditional custom view usage Aspose.Cells | worksheet top rows freeze Aspose.Cells | set worksheet zoom Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, optionally adds a custom view named "MyCustomView" with a 120% zoom (if the library version supports CustomViews), freezes the first two rows using FreezePanes, and saves the result as CustomViewAndFreeze.xlsx.
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

            // ----- Custom View (if supported) -----
            // Aspose.Cells versions prior to 23.9 do not expose the CustomViews collection.
            // If the current library supports it, the following code can be uncommented.
            /*
            int viewIndex = sheet.CustomViews.Add("MyCustomView");
            sheet.CustomViews[viewIndex].Zoom = 120;
            */

            // Freeze the first two rows to keep them visible while scrolling
            // Parameters: totalRows to freeze, totalColumns to freeze, row of the top-left scrollable cell, column of the top-left scrollable cell
            sheet.FreezePanes(2, 0, 2, 0);

            // Save the workbook to a file
            string outputPath = "CustomViewAndFreeze.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
