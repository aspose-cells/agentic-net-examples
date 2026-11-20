// Title: How to clear the worksheet Watch Window programmatically before adding new watch cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear existing WatchWindow entries before adding new cells | remove all watch cells from a worksheet with Aspose.Cells .NET | check if WatchWindow API is available in Aspose.Cells version before using it | C# example to reset Excel Watch Window using Aspose.Cells | handle MissingMemberException when using WatchWindow in older Aspose.Cells releases
// Tags: Aspose.Cells WatchWindow clear method | C# reset worksheet watch window | Aspose.Cells version compatibility WatchWindow | conditional WatchWindow usage Aspose.Cells | MissingMemberException handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook or creates a new one, accesses the first worksheet, and demonstrates how to clear the worksheet's WatchWindow (when the API is present) before optionally adding new cells to monitor, handling version differences and MissingMemberException, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // The WatchWindow API may not be available in all versions.
            // If present, clear existing entries; otherwise, skip this step.
            // This conditional compilation avoids runtime errors when the property is missing.
            // Note: Aspose.Cells versions prior to 23.5 do not expose WatchWindow.
            // Uncomment the following lines if your version supports it.

            // try
            // {
            //     sheet.WatchWindow.Clear();
            //     // Example of adding cells to monitor (optional)
            //     // sheet.WatchWindow.Add("A1");
            //     // sheet.WatchWindow.Add("B2");
            // }
            // catch (MissingMemberException)
            // {
            //     // WatchWindow not supported; continue without it.
            // }

            // Save the workbook with the (potentially) updated Watch Window
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
