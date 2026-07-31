// Title: C# – Update the "ReportPeriod" Named Range to C5:C15 with Aspose.Cells
// Description: This example shows how to load an Excel workbook with Aspose.Cells for .NET, locate the existing named range "ReportPeriod", determine whether it is global or sheet‑specific, change its RefersTo property to the absolute address =$C$5:$C$15 on the correct worksheet, and save the file.
// Keywords: Aspose.Cells | C# | named range update | RefersTo | C5:C15 | global named range | sheet‑specific name | Excel automation | Workbook.Names | Excel A1 notation
// Common Searches: Aspose.Cells change RefersTo C# | Update existing named range Aspose.Cells | Set named range to C5:C15 in .NET | Handle global vs sheet specific named ranges Aspose | C# code to modify Excel named range
// Developer Intent: Modify the "ReportPeriod" named range so it points to cells C5:C15.
// Use Cases: Fix a named range after inserting rows that shift the original range. | Resize a reporting period range when new data rows are added. | Align a named range with a newly created data block without recreating the workbook. | Standardize named ranges across multiple worksheets in a template.
// AI Prompts: Generate C# code using Aspose.Cells that finds a named range, checks for its existence, and updates its RefersTo to =$C$5:$C$15, handling both global and sheet‑specific scopes. | Explain step‑by‑step how to safely change the reference of an existing named range in an Excel file with Aspose.Cells. | Create a reusable C# method that accepts a workbook path, a named range name, and a new address, then updates the range's RefersTo property accordingly.

using System;
using Aspose.Cells;

// This example shows how to load an Excel workbook with Aspose.Cells for .NET, locate the existing named range "ReportPeriod", determine whether it is global or sheet‑specific, change its RefersTo property to the absolute address =$C$5:$C$15 on the correct worksheet, and save the file.
class UpdateNamedRange
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range "ReportPeriod"
        // The Names collection can be accessed by name directly
        Name reportPeriod = workbook.Worksheets.Names["ReportPeriod"];

        // Ensure the named range exists
        if (reportPeriod == null)
        {
            Console.WriteLine("Named range 'ReportPeriod' not found.");
            return;
        }

        // Determine the worksheet the name belongs to.
        // If SheetIndex == 0 the name is global; otherwise it refers to a specific sheet (1‑based index).
        Worksheet targetSheet = reportPeriod.SheetIndex == 0
            ? workbook.Worksheets[0]                     // default to first sheet for global names
            : workbook.Worksheets[reportPeriod.SheetIndex - 1];

        // Update the reference to span cells C5:C15 on the target worksheet
        // RefersTo must start with an equal sign and use absolute A1 notation.
        reportPeriod.RefersTo = $"={targetSheet.Name}!$C$5:$C$15";

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");

        Console.WriteLine("Named range 'ReportPeriod' updated successfully.");
    }
}
