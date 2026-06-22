using System;
using Aspose.Cells;

namespace RibbonButtonDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define custom Ribbon XML with a button that would invoke a macro named "LogButtonAction"
                string ribbonXml = @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
                                        <ribbon>
                                            <tabs>
                                                <tab id=""customTab"" label=""Debug Tab"">
                                                    <group id=""debugGroup"" label=""Debug Group"">
                                                        <button id=""logButton"" label=""Log Action"" size=""large"" onAction=""LogButtonAction""/>
                                                    </group>
                                                </tab>
                                            </tabs>
                                        </ribbon>
                                    </customUI>";

                // Assign the Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                Console.WriteLine("Ribbon button 'Log Action' has been added to the workbook.");

                // Save the workbook as a macro‑enabled file
                string filePath = "RibbonButtonDemo.xlsm";
                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved to {filePath}");

                // Simulate the button click action
                SimulateButtonAction(workbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simulated logic for ribbon button press
        static void SimulateButtonAction(Workbook workbook)
        {
            try
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue($"Button clicked at {DateTime.Now}");
                Console.WriteLine("Ribbon button action executed: cell A1 updated with timestamp.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Button action error: {ex.Message}");
            }
        }
    }
}