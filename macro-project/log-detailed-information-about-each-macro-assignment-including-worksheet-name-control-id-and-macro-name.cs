// Title: Log macro assignments of button controls in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, finds Button shapes, reads their Macro property via dynamic binding, and writes worksheet name, button name, and macro name to a text file. | Create a C# routine that expands the macro logger to also include ComboBox controls and outputs the results in CSV format using Aspose.Cells. | Write C# error‑handling logic for an Aspose.Cells macro logger that skips shapes without a Macro property while still recording successful button assignments.
// Common Searches: how to retrieve macro name from Excel button using Aspose.Cells C# | Aspose.Cells log button control macro assignments to file | C# iterate worksheet shapes and get Macro property with Aspose.Cells | save Excel macro mapping (worksheet, control, macro) using Aspose.Cells .NET | dynamic access to Macro property for button shapes in Aspose.Cells
// Tags: Aspose.Cells retrieve button macro property | C# log Excel control assignments | write macro mapping to text file Aspose.Cells | dynamic macro access Aspose.Cells .NET | error handling shape processing Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its shapes, identifies Button controls, reads the assigned Macro via dynamic binding, and writes a log entry containing the worksheet name, button name (control ID), and macro name to a text file while handling shape‑specific errors.
class MacroLogger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string logPath = "MacroAssignmentsLog.txt";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a log file to store macro assignment details
            using (StreamWriter logWriter = new StreamWriter(logPath))
            {
                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each shape (control) on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Check if the shape is a Button control
                            if (shape is Button)
                            {
                                // Use dynamic to access the Macro property (avoids compile‑time binding issues)
                                dynamic btn = shape;
                                string macro = btn.Macro as string;

                                if (!string.IsNullOrEmpty(macro))
                                {
                                    string logEntry = $"Worksheet: {sheet.Name}, Control ID: {btn.Name}, Macro: {macro}";
                                    Console.WriteLine(logEntry);
                                    logWriter.WriteLine(logEntry);
                                }
                            }
                        }
                        catch (Exception exShape)
                        {
                            // Log shape‑specific errors without stopping the whole process
                            Console.WriteLine($"Error processing shape \"{shape.Name}\": {exShape.Message}");
                        }
                    }
                }
            }

            Console.WriteLine($"Macro assignment log has been saved to \"{logPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
