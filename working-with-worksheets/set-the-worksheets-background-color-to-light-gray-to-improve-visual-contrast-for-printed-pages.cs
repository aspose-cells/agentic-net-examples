// Title: Setting a worksheet’s printed page background to light gray with Aspose.Cells for .NET – handling the missing API property
// Common Searches: Aspose.Cells C# set worksheet page background color for printing | How to change printed page shade of an Excel sheet using Aspose.Cells .NET | Is there a PrintPageBackgroundColor property in Aspose.Cells PageSetup | Workaround for missing worksheet print background color in Aspose.Cells | C# Aspose.Cells example to adjust Excel print page background color
// Tags: Aspose.Cells worksheet print background color | C# set Excel printed page shade | Aspose.Cells PageSetup background color limitation | Excel print page background color Aspose.Cells | workaround missing PrintPageBackgroundColor Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, and highlights that the current Aspose.Cells .NET API does not expose a PrintPageBackground or PrintPageBackgroundColor property in the PageSetup class, meaning the worksheet’s printed page background color (e.g., light gray) must be set manually after generation or via a workaround. The workbook is then saved as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // NOTE: The PageSetup class in the current Aspose.Cells version does not expose
            // PrintPageBackground or PrintPageBackgroundColor properties. If needed,
            // these settings can be adjusted via Excel UI after the file is generated.

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
