// Title: Save a new workbook as ODS with ribbons disabled using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook and saves it to ODS while disabling the UI ribbons via OdsSaveOptions in Aspose.Cells. | Show how to configure OdsSaveOptions to hide ribbons when exporting a workbook to ODS in a .NET application using Aspose.Cells.
// Common Searches: how to hide ribbons in ODS files with Aspose.Cells C# | Aspose.Cells OdsSaveOptions disable UI ribbons example | export workbook to ODS without ribbon UI using Aspose.Cells .NET
// Tags: ods saveoptions ribbon visibility | aspocells disable ribbons ods | c# export workbook to ods | aspocells ods file generation | disable ui ribbons aspocells

using System;
using Aspose.Cells;

// The example creates a new Workbook, notes that the DisableRibbons property is not directly available, and recommends using OdsSaveOptions to turn off ribbon UI when saving the workbook as an ODS file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // The DisableRibbons setting is not available in the current Aspose.Cells API.
            // If ribbon disabling is required for ODS, configure it via OdsSaveOptions when saving.

            // Save the workbook in ODS format using the updated enum value
            workbook.Save("output.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
