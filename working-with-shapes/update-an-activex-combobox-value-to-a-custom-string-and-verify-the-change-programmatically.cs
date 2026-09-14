// Title: Update an ActiveX ComboBox text in an Excel worksheet with Aspose.Cells for .NET and verify the change
// AI Prompts: Find the first ComboBox shape in the first worksheet, assign a custom string to its Text property, and save the workbook. | Read back the ComboBox.Text after setting it and output a success or failure message based on the comparison.
// Common Searches: how to set ActiveX ComboBox text using Aspose.Cells in C# | Aspose.Cells verify ComboBox value after modification | C# example for updating Excel form control ComboBox with custom string | retrieve and compare ComboBox.Text property in Aspose.Cells workbook | save Excel file after changing ActiveX ComboBox content with Aspose.Cells
// Tags: Aspose.Cells set ActiveX ComboBox Text | C# verify ComboBox value in Excel workbook | iterate worksheet shapes to locate ComboBox | save workbook after modifying form control | load workbook and update ComboBox content

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads 'ComboBoxDemo.xlsx', locates the first ActiveX ComboBox shape on the first worksheet, sets its Text property to 'MyCustomString', confirms the update succeeded, and saves the modified file as 'ComboBoxDemo_Updated.xlsx'.
class Program
{
    static void Main()
    {
        const string inputFile = "ComboBoxDemo.xlsx";
        const string outputFile = "ComboBoxDemo_Updated.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file '{inputFile}' not found.");
            return;
        }

        try
        {
            // Load the workbook that contains a ComboBox form control
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the ComboBox shape
            ComboBox comboBox = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape is ComboBox cb)
                {
                    comboBox = cb;
                    break;
                }
            }

            if (comboBox == null)
            {
                Console.WriteLine("ComboBox not found.");
                return;
            }

            // Set the ComboBox to a custom string value
            string customValue = "MyCustomString";
            comboBox.Text = customValue; // Update displayed text/value

            // Verify that the value was updated
            string currentValue = comboBox.Text;
            if (!string.IsNullOrEmpty(currentValue) && currentValue == customValue)
            {
                Console.WriteLine("ComboBox value successfully updated to: " + currentValue);
            }
            else
            {
                Console.WriteLine("Failed to update ComboBox value.");
            }

            // Save the workbook with the updated ComboBox value
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
