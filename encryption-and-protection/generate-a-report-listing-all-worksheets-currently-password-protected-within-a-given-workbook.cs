// Title: List password‑protected worksheets in an Excel workbook using Aspose.Cells for .NET
// Description: Load an Excel file with Aspose.Cells, iterate through its worksheets, detect those secured with a password via the Protection.IsProtectedWithPassword property, and output each sheet's name and index to the console without modifying the file.
// Keywords: Aspose.Cells password protected sheets | C# detect worksheet protection | list protected worksheets .NET | Excel sheet security check | enumerate protected worksheets
// Common Searches: how to find password protected worksheets with Aspose.Cells | C# list protected sheets in Excel file | detect worksheet password using Aspose.Cells for .NET | report protected worksheets without saving | enumerate Excel sheets that require a password
// Developer Intent: Retrieve and display the names and indexes of all worksheets that have password protection in a given workbook.
// Use Cases: Audit workbook security before distribution | Create a compliance log of protected sheets | Skip password‑protected tabs during data extraction
// AI Prompts: Generate C# code with Aspose.Cells that writes the list of password‑protected worksheets to a text file. | Show how to log protected worksheet names using a logging framework instead of Console.WriteLine. | Extend the example to also report the specific protection options (e.g., objects, scenarios) enabled on each sheet.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells, iterate through its worksheets, detect those secured with a password via the Protection.IsProtectedWithPassword property, and output each sheet's name and index to the console without modifying the file.
class ListPasswordProtectedWorksheets
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all worksheets and list those protected with a password
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Check if the worksheet's protection has a password
            if (sheet.Protection.IsProtectedWithPassword)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' (index {i}) is password protected.");
            }
        }

        // No need to save the workbook for this reporting task
    }
}
