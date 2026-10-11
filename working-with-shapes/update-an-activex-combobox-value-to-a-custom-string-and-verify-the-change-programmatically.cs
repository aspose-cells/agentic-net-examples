// Title: Set and verify a custom string for an ActiveX ComboBox in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that finds an ActiveX ComboBox named 'ComboBox1' on a worksheet, assigns a custom string to its Value property via OleObject, confirms the assignment, and saves the workbook. | Show how to use the dynamic keyword to access the OleFormat and underlying COM object of an OleObject for modifying an ActiveX ComboBox text in Aspose.Cells. | Adapt the sample so the new ComboBox value is read from a specific cell (e.g., A1) before updating the control.
// Common Searches: aspocells set activeX combobox value c# example | c# update oleobject combobox text in excel workbook | verify activeX control value after modification using Aspose.Cells | dynamic OleFormat access for ActiveX controls in Aspose.Cells .NET | save workbook after changing ActiveX ComboBox with Aspose.Cells
// Tags: Aspose.Cells ActiveX ComboBox value assignment | C# dynamic OleObject manipulation | Excel ActiveX control update via OleFormat | programmatic verification of ComboBox text | saving workbook after ActiveX modification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code loads an existing Excel file, locates an ActiveX ComboBox named 'ComboBox1' through the worksheet's OleObjects collection, uses dynamic casting to reach the OleFormat and underlying COM object, sets its Value to a custom string, verifies the assignment by comparing the returned value, and finally saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the ActiveX ComboBox
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the ActiveX ComboBox by its name using the OleObjects collection
            OleObject oleObject = null;
            foreach (OleObject obj in sheet.OleObjects)
            {
                if (obj.Name == "ComboBox1")
                {
                    oleObject = obj;
                    break;
                }
            }

            if (oleObject != null)
            {
                const string customText = "My Custom String";

                try
                {
                    // Use dynamic to access OleFormat and the underlying COM object
                    dynamic dynOleObject = oleObject;
                    dynamic oleFormat = dynOleObject.OleFormat;   // runtime member access
                    dynamic comObj = oleFormat.Object;            // underlying ActiveX object

                    // Set the ComboBox value
                    comObj.Value = customText;

                    // Verify that the value was set correctly
                    string currentValue = comObj.Value?.ToString();
                    if (currentValue == customText)
                    {
                        Console.WriteLine("ComboBox value updated successfully.");
                    }
                    else
                    {
                        Console.WriteLine("ComboBox value update failed.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while updating ComboBox value: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("ActiveX ComboBox named 'ComboBox1' was not found.");
            }

            // Save the workbook with the updated ComboBox value
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
