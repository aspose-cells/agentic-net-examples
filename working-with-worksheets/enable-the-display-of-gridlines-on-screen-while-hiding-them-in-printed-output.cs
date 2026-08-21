// Title: Display screen gridlines while suppressing print gridlines using Aspose.Cells for .NET (C#)
// Description: This C# example creates a new workbook, turns on worksheet gridlines for on‑screen viewing (Worksheet.IsGridlinesVisible = true) and disables them for printed output (Worksheet.PageSetup.PrintGridlines = false). Sample text is added and the workbook is saved as GridlinesScreenOnly.xlsx.
// Keywords: Aspose.Cells C# gridlines | Worksheet.IsGridlinesVisible | PageSetup.PrintGridlines | hide gridlines in print | show gridlines on screen | Excel gridline visibility .NET | Aspose.Cells print settings | gridlines only on screen
// Common Searches: Aspose.Cells show gridlines on screen only | hide gridlines when printing Excel with Aspose.Cells | Worksheet.IsGridlinesVisible property example | PageSetup.PrintGridlines false C# | C# Aspose.Cells gridline visibility tutorial
// Developer Intent: Enable gridlines for workbook display while ensuring they are omitted from printed output.
// Use Cases: Interactive Excel reports where gridlines aid data entry but printed copies need a clean layout. | Templates for internal users that show cell boundaries on screen, yet produce professional PDFs without gridlines. | Dashboards that rely on visual gridlines for navigation in the application but require a polished look in hard‑copy distribution.
// AI Prompts: Generate C# code with Aspose.Cells that shows gridlines on screen and hides them when printing, then saves the file as .xlsx. | Explain the interaction between Worksheet.IsGridlinesVisible and PageSetup.PrintGridlines for controlling screen versus print gridline visibility in Aspose.Cells. | Provide a step‑by‑step guide to toggle gridline visibility for on‑screen view and print output in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesDemo
{
    // This C# example creates a new workbook, turns on worksheet gridlines for on‑screen viewing (Worksheet.IsGridlinesVisible = true) and disables them for printed output (Worksheet.PageSetup.PrintGridlines = false). Sample text is added and the workbook is saved as GridlinesScreenOnly.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Show gridlines on screen
            worksheet.IsGridlinesVisible = true;

            // Hide gridlines when printing
            worksheet.PageSetup.PrintGridlines = false;

            // Add some sample data to visualize the gridlines
            worksheet.Cells["A1"].PutValue("Gridlines visible on screen");
            worksheet.Cells["A2"].PutValue("but not printed.");

            // Save the workbook
            workbook.Save("GridlinesScreenOnly.xlsx");

            Console.WriteLine("Workbook saved with screen gridlines visible and print gridlines hidden.");
        }
    }
}
