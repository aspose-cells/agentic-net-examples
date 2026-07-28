// Title: Find Hidden Protected Worksheets in an Excel File with Aspose.Cells for .NET
// Description: Creates a workbook, hides and protects a sheet, saves it, then reloads the file and scans all worksheets using the IsVisible and IsProtected properties to list any sheets that are both hidden and protected.
// Keywords: Aspose.Cells hidden worksheet detection | protected sheet list C# | Excel hidden protected sheet .NET | Aspose.Cells IsVisible | Aspose.Cells IsProtected | C# Excel security audit | Aspose.Cells sample code
// Common Searches: Aspose.Cells detect hidden sheets | list protected worksheets C# | how to check if worksheet is hidden and protected Aspose.Cells | C# find very hidden worksheets Aspose.Cells | Excel workbook security audit Aspose.Cells
// Developer Intent: Locate worksheets that are simultaneously hidden (or VeryHidden) and protected, and retrieve their names for review or reporting.
// Use Cases: Perform a security audit before distributing a workbook to ensure no confidential hidden sheets remain. | Integrate a compliance check into CI/CD pipelines that flags hidden protected worksheets. | Generate a report of all hidden protected sheets for governance or documentation purposes. | Automate data‑governance scans across multiple Excel files in an enterprise environment.
// AI Prompts: Write a C# method that returns a List<string> of worksheet names that are hidden and protected using Aspose.Cells. | Provide code to log the detected hidden protected worksheet names to a timestamped text file. | Show how to extend the detection loop to include VeryHidden sheets and optionally verify a protection password. | Create a PowerShell script that invokes the Aspose.Cells .NET assembly to list hidden protected worksheets in a given Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsHiddenProtectedDemo
{
    // Creates a workbook, hides and protects a sheet, saves it, then reloads the file and scans all worksheets using the IsVisible and IsProtected properties to list any sheets that are both hidden and protected.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook with a hidden and protected worksheet
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create
            // Add two more worksheets
            workbook.Worksheets.Add("VisibleSheet");
            workbook.Worksheets.Add("HiddenProtectedSheet");

            // Populate data (optional)
            workbook.Worksheets["VisibleSheet"].Cells["A1"].PutValue("Visible data");
            workbook.Worksheets["HiddenProtectedSheet"].Cells["A1"].PutValue("Secret data");

            // Hide the worksheet
            Worksheet hiddenSheet = workbook.Worksheets["HiddenProtectedSheet"];
            hiddenSheet.IsVisible = false;                         // hidden

            // Protect the same worksheet (without password for simplicity)
            hiddenSheet.Protect(ProtectionType.All);               // protected

            // Save the workbook to disk
            string filePath = "SampleWorkbook.xlsx";
            workbook.Save(filePath);                               // save

            // ---------------------------------------------------------------
            // 2. Load the workbook and detect hidden protected worksheets
            // ---------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);      // load

            bool anyHiddenProtected = false;
            Console.WriteLine("Hidden and protected worksheets found:");

            for (int i = 0; i < loadedWorkbook.Worksheets.Count; i++)
            {
                Worksheet ws = loadedWorkbook.Worksheets[i];

                // Worksheet.IsVisible == false indicates hidden (or VeryHidden)
                // Worksheet.IsProtected indicates protection status
                if (!ws.IsVisible && ws.IsProtected)
                {
                    anyHiddenProtected = true;
                    Console.WriteLine($"- {ws.Name}");
                }
            }

            if (!anyHiddenProtected)
            {
                Console.WriteLine("None.");
            }
        }
    }
}
