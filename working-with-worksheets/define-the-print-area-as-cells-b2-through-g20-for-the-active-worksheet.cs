// Title: Set the print area to B2:G20 on the active worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set print area B2:G20 on current worksheet | how to define a custom print range in Aspose.Cells .NET | programmatically set worksheet print area using Aspose.Cells PageSetup | C# Aspose.Cells example for setting print area before saving workbook | define print area for active sheet in Aspose.Cells for .NET
// Tags: Aspose.Cells set worksheet print area | PageSetup.PrintArea property C# | custom print area B2:G20 Aspose.Cells | active worksheet print area configuration .NET | save workbook after setting print area Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new Workbook, obtains the active worksheet, assigns the print area to the range B2:G20 via the PageSetup.PrintArea property, and then saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Load an existing workbook or create a new one
        // (Assuming a new workbook for demonstration)
        Workbook workbook = new Workbook();

        // Get the active worksheet (the first worksheet by default)
        Worksheet sheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Define the print area as B2:G20
        sheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
