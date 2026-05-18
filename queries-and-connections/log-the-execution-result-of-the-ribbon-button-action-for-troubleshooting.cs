using System;
using Aspose.Cells;

namespace RibbonButtonLoggingDemo
{
    // Simple class that represents the action performed when the ribbon button is clicked
    public class RibbonButtonHandler
    {
        // This method would be invoked by the ribbon button (via a macro or external call)
        public void ExecuteAction()
        {
            // Perform the desired operation – here we just simulate some work
            Console.WriteLine("Ribbon button action started.");

            // Simulated result
            string result = "Action completed successfully.";

            // Log the result
            Console.WriteLine($"Execution Result: {result}");
        }
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and define custom Ribbon XML
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(); // create a new workbook

            // Define a simple Ribbon UI with a button that calls a macro named "RunButtonAction"
            // (In a real Excel environment the macro would call RibbonButtonHandler.ExecuteAction)
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"Debug Tab\">" +
                "        <group id=\"debugGroup\" label=\"Debug Tools\">" +
                "          <button id=\"debugButton\" label=\"Run Action\" size=\"large\" onAction=\"RunButtonAction\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;   // from Aspose.Cells.Workbook.RibbonXml

            // ------------------------------------------------------------
            // 2. Simulate the ribbon button click by directly invoking the handler
            // ------------------------------------------------------------
            RibbonButtonHandler handler = new RibbonButtonHandler();
            handler.ExecuteAction();

            // ------------------------------------------------------------
            // 3. Save the workbook (the Ribbon UI will be persisted)
            // ------------------------------------------------------------
            workbook.Save("RibbonButtonDemo.xlsm");

            // Indicate that the process finished
            Console.WriteLine("Workbook saved with custom Ribbon. Check console output for logged execution result.");
        }
    }
}