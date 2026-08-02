using System;
using System.IO;
using Aspose.Cells;

namespace RibbonButtonLoggingDemo
{
    class Program
    {
        // Simulated action that would be triggered by the ribbon button
        static void ExecuteRibbonButtonAction()
        {
            try
            {
                // Core logic of the button (placeholder)
                int result = 42;

                // Log successful execution
                Console.WriteLine($"[INFO] Ribbon button action executed successfully. Result = {result}");
            }
            catch (Exception ex)
            {
                // Log any error that occurs during execution
                Console.WriteLine($"[ERROR] Ribbon button action failed: {ex.Message}");
            }
        }

        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define Ribbon XML with a custom button.
                // The button is linked to a macro named "LogButtonAction".
                // In a real Excel environment the macro would call ExecuteRibbonButtonAction.
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"Debug Tab\">" +
                    "        <group id=\"debugGroup\" label=\"Debug Tools\">" +
                    "          <button id=\"logButton\" label=\"Run Action\" size=\"large\" onAction=\"LogButtonAction\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Assign the Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                // Define output file path
                string outputPath = "RibbonButtonDemo.xlsm";

                // Save the workbook as a macro‑enabled file (XLSM)
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");

                // Simulate the button click for troubleshooting purposes
                Console.WriteLine("Simulating ribbon button click...");
                ExecuteRibbonButtonAction();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"[ERROR] Unexpected error: {ex.Message}");
            }
            finally
            {
                // Keep console open
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }
    }
}