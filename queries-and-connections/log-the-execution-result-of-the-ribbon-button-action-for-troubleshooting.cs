// Title: How to log the execution result of a custom Ribbon button in an Aspose.Cells macro‑enabled workbook using C#
// AI Prompts: Write C# code that creates a macro‑enabled workbook with a custom Ribbon XML button and logs the button's action result to a text file. | Add robust exception handling to the ribbon button handler so that both success and error messages are appended to a persistent log. | Generate a C# method that reads the latest entry from RibbonActionLog.txt and displays the timestamp of the most recent button execution.
// Common Searches: C# log custom UI button click in Aspose.Cells macro workbook | Aspose.Cells add Ribbon XML button and write execution details to file | how to capture and persist ribbon button action result in Excel using Aspose.Cells | troubleshooting Aspose.Cells custom UI button by logging output to text file | save macro‑enabled .xlsm with custom ribbon and log button activity C#
// Tags: ribbon UI button logging Aspose.Cells | macro-enabled workbook Ribbon XML C# | append execution result to text file | exception handling for ribbon button action | persisted log for Excel custom UI

using System;
using System.IO;
using Aspose.Cells;

// The example creates a macro‑enabled .xlsm workbook with a custom Ribbon XML button using Aspose.Cells, saves it, and then simulates the button click. The button handler writes a success message with a timestamp to both the console and a persistent text log, while also capturing and logging any exceptions.
class RibbonButtonLogger
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML with a button that calls a macro named "OnCustomButtonClick"
            string ribbonXml =
                @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
                    <ribbon>
                        <tabs>
                            <tab id=""customTab"" label=""My Tab"">
                                <group id=""customGroup"" label=""My Group"">
                                    <button id=""customButton"" label=""Run Action"" size=""large"" onAction=""OnCustomButtonClick""/>
                                </group>
                            </tab>
                        </tabs>
                    </ribbon>
                  </customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file (the button will appear when opened in Excel)
            string filePath = "RibbonDemo.xlsm";
            workbook.Save(filePath, SaveFormat.Xlsm);

            Console.WriteLine("Workbook with custom Ribbon saved to " + filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("[ERROR] Failed to create or save workbook: " + ex.Message);
        }

        // Simulate the ribbon button click for troubleshooting purposes
        ExecuteRibbonButtonAction();
    }

    // This method represents the action that would be triggered by the ribbon button.
    // It logs the execution result to both console and a simple text log file.
    static void ExecuteRibbonButtonAction()
    {
        try
        {
            // Example operation performed by the button
            string result = $"Action executed successfully at {DateTime.Now:O}";

            // Log to console (Info level)
            Console.WriteLine("[INFO] Ribbon button action result: " + result);

            // Append the result to a log file for persistent troubleshooting
            File.AppendAllText("RibbonActionLog.txt", result + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // Log any errors (Error level)
            Console.WriteLine("[ERROR] Ribbon button action failed: " + ex.Message);
            File.AppendAllText("RibbonActionLog.txt", "Error: " + ex.Message + Environment.NewLine);
        }
    }
}
