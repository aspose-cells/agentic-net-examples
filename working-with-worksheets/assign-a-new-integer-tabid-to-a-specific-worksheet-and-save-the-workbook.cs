// Title: Assign a custom integer TabId to a worksheet and save the workbook using Aspose.Cells for .NET
// AI Prompts: Load an Excel file (or create a new workbook if it does not exist), locate a worksheet by name, set its TabId property to a given integer, and write the workbook to a new file with Aspose.Cells in C#. | Using Aspose.Cells for .NET, retrieve a specific sheet (or the first sheet), assign a numeric TabId value, and persist the change by saving the workbook to a different path.
// Common Searches: C# Aspose.Cells set worksheet TabId example | How to change Excel sheet tab identifier programmatically with Aspose.Cells | Saving workbook after modifying TabId property in .NET | Assign numeric TabId to a specific worksheet using Aspose.Cells API
// Tags: Aspose.Cells worksheet TabId assignment | C# Excel TabId property update | Aspose.Cells workbook save after TabId change | Excel sheet TabId manipulation .NET | load or create workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing workbook or creates a new one, obtains the worksheet named "Sheet1" (or the first sheet), assigns the integer 12345 to its TabId property, and saves the updated workbook to a separate output file.
class AssignTabIdExample
{
    static void Main()
    {
        try
        {
            // Define the folder and file names.
            string dataDir = @"C:\Data\";
            string inputFile = Path.Combine(dataDir, "input.xlsx");
            string outputFile = Path.Combine(dataDir, "output.xlsx");

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing.
            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default sheet
                workbook.Worksheets[0].Name = "Sheet1";
                // Optionally save the newly created workbook for future runs.
                workbook.Save(inputFile);
            }

            // Get the worksheet named "Sheet1" or fallback to the first worksheet.
            Worksheet worksheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            // Assign a new integer TabId to the worksheet.
            worksheet.TabId = 12345;

            // Save the workbook with the updated TabId.
            workbook.Save(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
