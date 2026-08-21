// Title: Enable printing of gridlines in an Excel worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, sets worksheet.IsGridlinesVisible to true, enables PageSetup.PrintGridlines, adds sample data, and saves the file so that gridlines appear both on screen and in the printed output.
// Keywords: Aspose.Cells | C# | PrintGridlines | IsGridlinesVisible | Excel gridlines | worksheet printing | PageSetup | export to Excel | gridline visibility | Aspose.Cells .NET
// Common Searches: Aspose.Cells print gridlines C# | How to show gridlines when printing Excel with Aspose | Enable worksheet gridlines in Aspose.Cells | Set PrintGridlines property Aspose.Cells | Make gridlines visible on screen Aspose.Cells
// Developer Intent: Configure a worksheet so gridlines appear on screen and are included in printed output.
// Use Cases: Generate reports that retain cell borders when printed for easier data verification. | Create invoices or statements where gridlines improve readability on paper copies. | Automate export of data tables where alignment must be preserved in both screen view and hard copy.
// AI Prompts: Write C# code with Aspose.Cells that turns on IsGridlinesVisible and PrintGridlines, then saves the workbook as PDF. | Show how to toggle the PrintGridlines setting based on a user‑defined boolean flag in Aspose.Cells. | Provide an example that applies IsGridlinesVisible and PrintGridlines to all worksheets in a multi‑sheet workbook.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets worksheet.IsGridlinesVisible to true, enables PageSetup.PrintGridlines, adds sample data, and saves the file so that gridlines appear both on screen and in the printed output.
public class EnablePrintGridlines
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data so the gridlines are visible in the output
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B2"].PutValue(123);

            // Make gridlines visible on screen (optional, but often desired)
            worksheet.IsGridlinesVisible = true;

            // Enable printing of gridlines
            worksheet.PageSetup.PrintGridlines = true;

            // Save the workbook with the gridlines setting applied
            workbook.Save("PrintGridlinesEnabled.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        EnablePrintGridlines.Run();
    }
}
