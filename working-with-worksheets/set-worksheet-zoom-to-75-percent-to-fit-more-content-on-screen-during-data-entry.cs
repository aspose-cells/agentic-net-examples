// Title: C# – Aspose.Cells: Set Worksheet Zoom to 75 %
// Description: Shows how to use Aspose.Cells for .NET to set the Zoom property of a worksheet to 75 % and save the workbook.
// Keywords: Aspose.Cells Worksheet.Zoom | C# Excel zoom level | set worksheet zoom 75 percent | adjust Excel view programmatically | Aspose.Cells .NET zoom factor | Excel zoom API US | Excel zoom API EU
// Common Searches: how to set worksheet zoom to 75% with Aspose.Cells C# | Aspose.Cells change Excel sheet zoom programmatically | C# set zoom factor for Excel worksheet using Aspose.Cells | adjust worksheet view percentage Aspose.Cells .NET | set Excel zoom for data entry Aspose.Cells example
// Developer Intent: Apply a 75 % zoom setting to a worksheet via Aspose.Cells.
// Use Cases: Create a new workbook where the first sheet opens at 75 % zoom to display more columns during data entry. | Open an existing report, modify the zoom of a target worksheet to 75 % for consistent on‑screen layout across devices, then save. | Generate a template workbook with a predefined 75 % zoom so end users do not need to adjust view settings manually.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, sets the first worksheet's zoom to 75 %, and saves the changes. | Explain how to apply different zoom percentages to multiple worksheets in a single workbook using Aspose.Cells for .NET. | Provide a step‑by‑step guide for programmatically changing the zoom level of a worksheet to improve data‑entry visibility with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to set the Zoom property of a worksheet to 75 % and save the workbook.
public class SetWorksheetZoom
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the zoom factor to 75%
        worksheet.Zoom = 75;

        // Save the workbook to a file
        workbook.Save("WorksheetZoom75.xlsx");
        Console.WriteLine("Workbook saved as WorksheetZoom75.xlsx");
    }
}
