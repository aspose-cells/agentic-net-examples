// Title: Check for an ActiveX ComboBox on the second worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, accesses the worksheet at index 1, and determines whether an OleObject named "ComboBox1" exists. | Show how to iterate over all OleObjects on a given worksheet in Aspose.Cells and provide a fallback when the control name is unknown. | Explain how to safely handle a missing workbook file and output a clear message about the presence of an ActiveX ComboBox using Aspose.Cells.
// Common Searches: Aspose.Cells C# find ActiveX ComboBox on second worksheet | how to verify OleObject named ComboBox1 in Excel with Aspose.Cells | enumerate ActiveX controls on a specific sheet using Aspose.Cells .NET | detect presence of ActiveX ComboBox in .xlsx file via C# | Aspose.Cells cannot read ComboBox value but can check control existence
// Tags: enumerate oleobjects Aspose.Cells C# | detect activeX combobox presence Aspose.Cells | access second worksheet oleobject Aspose.Cells | handle missing excel file Aspose.Cells | activeX oleobject lookup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// The example loads an Excel workbook, accesses the second worksheet, searches for an ActiveX ComboBox OleObject named "ComboBox1" (or falls back to the first OleObject), notes that Aspose.Cells does not expose the selected value, and prints a message indicating whether the control was found.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (index 1)
            Worksheet sheet = workbook.Worksheets[1];

            // Locate the ActiveX ComboBox OleObject.
            // If you know the name of the control, replace "ComboBox1" with that name.
            OleObject comboOle = null;
            foreach (OleObject ole in sheet.OleObjects)
            {
                if (ole.Name == "ComboBox1")
                {
                    comboOle = ole;
                    break;
                }
            }

            // If the control name is unknown, fall back to the first OleObject on the sheet.
            if (comboOle == null && sheet.OleObjects.Count > 0)
            {
                comboOle = sheet.OleObjects[0];
            }

            // Aspose.Cells does not expose ActiveX control values directly,
            // so we simply indicate whether the control was found.
            string selectedValue = comboOle != null
                ? "Control found (value retrieval not supported)"
                : "Control not found";

            // Output the result.
            Console.WriteLine("Selected value: " + selectedValue);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
