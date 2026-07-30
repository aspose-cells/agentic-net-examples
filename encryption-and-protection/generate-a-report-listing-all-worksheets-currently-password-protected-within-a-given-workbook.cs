// Title: List password‑protected worksheets in an Excel workbook with Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel file using Aspose.Cells, scans all worksheets, detects those protected with a password via Protection.IsProtectedWithPassword, and prints their indexes and names while handling missing files and disposing resources.
// Keywords: Aspose.Cells list protected worksheets | C# detect worksheet password protection | enumerate protected sheets Aspose.Cells | check Excel worksheet protection .NET | worksheet protection status Aspose.Cells
// Common Searches: how to find password protected sheets using Aspose.Cells | Aspose.Cells code to list protected worksheets | C# check if Excel worksheet is password protected | enumerate protected worksheets Aspose.Cells .NET | detect worksheet protection without opening password
// Developer Intent: Retrieve the names and indexes of all worksheets in a workbook that have password protection enabled.
// Use Cases: Generate an audit report of protected sheets before sharing a workbook. | Automate compliance checks across multiple Excel files in a CI pipeline. | Trigger custom logic (e.g., notifications or de‑protection) for password‑protected worksheets during document processing.
// AI Prompts: Create C# code with Aspose.Cells that writes the names of password‑protected worksheets to a CSV file, including robust error handling. | Provide a method returning List<int> of worksheet indexes that are password protected, ensuring proper disposal of the Workbook object. | Show how to extend the sample to also retrieve and display any protection password hint for each protected worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel file using Aspose.Cells, scans all worksheets, detects those protected with a password via Protection.IsProtectedWithPassword, and prints their indexes and names while handling missing files and disposing resources.
    public class ListPasswordProtectedWorksheets
    {
        public static void Run()
        {
            // Path to the workbook to be inspected
            string workbookPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook (no password needed for opening)
                workbook = new Workbook(workbookPath);

                Console.WriteLine("Worksheets protected with a password:");
                bool anyProtected = false;

                // Iterate through all worksheets and collect those protected with a password
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Protection.IsProtectedWithPassword indicates password protection on the worksheet
                    if (sheet.Protection.IsProtectedWithPassword)
                    {
                        anyProtected = true;
                        Console.WriteLine($"- Sheet index {i}: \"{sheet.Name}\"");
                    }
                }

                if (!anyProtected)
                {
                    Console.WriteLine("None of the worksheets are password protected.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ListPasswordProtectedWorksheets.Run();
        }
    }
}
