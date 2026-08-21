// Title: Set Worksheet TabId with Aspose.Cells for .NET and Save the Workbook
// Description: Demonstrates how to assign a custom integer to the TabId property of a worksheet, save the workbook as an .xlsx file, reload it, and verify that the TabId value persists using Aspose.Cells for .NET.
// Keywords: Aspose.Cells TabId | set worksheet TabId .NET | Excel TabId property | save workbook Aspose | retrieve TabId after load | C# Aspose.Cells example
// Common Searches: how to set worksheet TabId using Aspose.Cells | Aspose.Cells save workbook with custom TabId | verify TabId value after reopening Excel file | C# example for TabId property Aspose
// Developer Intent: Assign a numeric TabId to a specific worksheet and ensure the value is stored in the saved Excel file.
// Use Cases: Tag worksheets with unique identifiers for internal tracking before export. | Maintain worksheet IDs across sessions to map Excel sheets to database records. | Enable downstream processes to locate worksheets by their persisted TabId.
// AI Prompts: Write C# code that sets the TabId of the second worksheet to 9876, saves as 'Report.xlsx', and reads the value back for verification. | Show how to loop through all worksheets in a loaded workbook and print each worksheet's TabId using Aspose.Cells for .NET. | Explain best practices for using the TabId property to link Excel worksheets with external metadata when saving and loading files.

using System;
using Aspose.Cells;

// Demonstrates how to assign a custom integer to the TabId property of a worksheet, save the workbook as an .xlsx file, reload it, and verify that the TabId value persists using Aspose.Cells for .NET.
class SetWorksheetTabId
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Assign a new TabId value
        worksheet.TabId = 12345;

        // Save the workbook to a file
        string filePath = "TabIdDemo.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook to verify the TabId
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
        Console.WriteLine("Loaded Worksheet TabId: " + loadedWorksheet.TabId);
    }
}
