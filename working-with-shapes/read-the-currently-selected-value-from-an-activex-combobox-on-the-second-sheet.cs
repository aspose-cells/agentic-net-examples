// Title: Read the selected value of an ActiveX ComboBox (ComboBox1) from the second worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel workbook, navigates to the second sheet, locates the OleObject named "ComboBox1", and returns its current selection as a string. | Show how to obtain the ObjectData byte array of an ActiveX ComboBox on a specific worksheet with Aspose.Cells and convert the bytes to readable UTF‑8 text.
// Common Searches: Aspose.Cells get value from ActiveX ComboBox on specific sheet C# | How to read OleObject ObjectData of a ComboBox in Excel using .NET | Retrieve selected item of ActiveX ComboBox from second worksheet with Aspose.Cells | C# example reading ActiveX control data from workbook using Aspose.Cells | Extract ComboBox selected text from Excel file programmatically
// Tags: Aspose.Cells read ActiveX OleObject data | C# read ComboBox selected text | second worksheet OleObject handling Aspose.Cells | decode ObjectData UTF-8 Aspose.Cells | load workbook access ActiveX controls .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads 'input.xlsx', checks for a second worksheet, searches its OleObjects for an ActiveX ComboBox named 'ComboBox1', reads the control's raw ObjectData byte array, decodes it as UTF‑8, and prints the selected value.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Ensure the second worksheet exists (zero‑based index).
            if (workbook.Worksheets.Count <= 1)
            {
                Console.WriteLine("Error: The workbook does not contain a second worksheet.");
                return;
            }

            Worksheet sheet = workbook.Worksheets[1];

            // Find the OleObject (ActiveX ComboBox) named "ComboBox1".
            OleObject comboOle = null;
            foreach (OleObject ole in sheet.OleObjects)
            {
                if (string.Equals(ole.Name, "ComboBox1", StringComparison.OrdinalIgnoreCase))
                {
                    comboOle = ole;
                    break;
                }
            }

            if (comboOle != null)
            {
                // Obtain the raw binary data of the control.
                byte[] rawData = comboOle.ObjectData;

                // For demonstration, interpret the binary data as UTF‑8 text.
                string selectedValue = System.Text.Encoding.UTF8.GetString(rawData);

                Console.WriteLine("Selected value (raw data interpreted as UTF‑8): " + selectedValue);
            }
            else
            {
                Console.WriteLine("ActiveX ComboBox named 'ComboBox1' was not found on the second sheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
