using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class ReadActiveXComboBoxValue
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (zero‑based index)
            Worksheet sheet = workbook.Worksheets[1];

            // Iterate through all shapes on the sheet to locate ActiveX ComboBox controls
            foreach (Shape shape in sheet.Shapes)
            {
                // Ensure the shape hosts an ActiveX control and that it is a ComboBox
                if (shape.ActiveXControl is ComboBoxActiveXControl comboBox)
                {
                    // Read the currently selected value (the Value property reflects the selected item)
                    string selectedValue = comboBox.Value;

                    // Output the value
                    Console.WriteLine("Selected value of ActiveX ComboBox: " + selectedValue);
                }
            }

            // (Optional) Save the workbook if any changes were made
            // workbook.Save("OutputWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}