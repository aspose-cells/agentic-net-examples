// Title: Export Formula Audit Trail with Timestamps to XML using Aspose.Cells (C#)
// Description: Loads an Excel workbook, switches to manual calculation mode, and attaches a custom AbstractCalculationMonitor that records sheet index, row, column, and UTC timestamp before each formula is evaluated. The collected entries are assembled into an XML document (FormulaAuditTrail) and saved as a report file.
// Keywords: Aspose.Cells | C# formula audit trail | Excel calculation log | XML report | custom calculation monitor | record timestamps | manual calculation mode | export audit XML | formula audit | Aspose.Cells example
// Common Searches: Aspose.Cells export formula audit trail | C# log Excel formula calculations with timestamps | create calculation monitor in Aspose.Cells | generate XML audit of Excel formulas | save formula evaluation order to XML Aspose.Cells
// Developer Intent: Create an XML report that logs every formula calculation and its timestamp from an Excel workbook using Aspose.Cells.
// Use Cases: Debug complex spreadsheets by reviewing calculation order and timing. | Provide compliance documentation of calculation timestamps for regulated financial models. | Monitor performance of workbook recalculation in automated pipelines. | Integrate the audit XML into CI/CD validation or monitoring systems.
// AI Prompts: Write C# code that implements an Aspose.Cells AbstractCalculationMonitor to capture sheet, row, column, and UTC timestamp for each formula evaluation and output an XML file. | Explain how to set an Aspose.Cells workbook to manual calculation mode and attach a custom calculation monitor for auditing purposes. | Show how to extend the AuditCalculationMonitor to record the duration of each cell calculation and include it in the XML audit trail. | Generate a GitHub‑compatible README snippet describing this example, its prerequisites, and how to run it.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

namespace FormulaAuditTrailExport
{
    // Custom calculation monitor that records each cell calculation with a timestamp
    // Loads an Excel workbook, switches to manual calculation mode, and attaches a custom AbstractCalculationMonitor that records sheet index, row, column, and UTC timestamp before each formula is evaluated. The collected entries are assembled into an XML document (FormulaAuditTrail) and saved as a report file.
    public class AuditCalculationMonitor : AbstractCalculationMonitor
    {
        // List to hold audit entries
        private readonly List<XElement> _entries = new List<XElement>();

        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Capture the current UTC time
            string timestamp = DateTime.UtcNow.ToString("o"); // ISO 8601 format

            // Build an XML element for this calculation step
            XElement entry = new XElement("Calculation",
                new XAttribute("SheetIndex", sheetIndex),
                new XAttribute("Row", rowIndex),
                new XAttribute("Column", columnIndex),
                new XAttribute("Timestamp", timestamp));

            _entries.Add(entry);
        }

        // Called after a cell is calculated (optional, not used here)
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex) { }

        // Called when a circular reference is detected (optional, not used here)
        public override bool OnCircular(IEnumerator circularCellsData) => base.OnCircular(circularCellsData);

        // Generates the final XML document containing all audit entries
        public XDocument GetAuditXml()
        {
            return new XDocument(
                new XElement("FormulaAuditTrail",
                    new XAttribute("GeneratedOn", DateTime.UtcNow.ToString("o")),
                    _entries));
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "InputWorkbook.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure manual calculation mode so we control when formulas are evaluated
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Create the custom monitor
                AuditCalculationMonitor monitor = new AuditCalculationMonitor();

                // Set up calculation options and attach the monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = monitor
                };

                // Perform calculation; the monitor will record each cell's calculation timestamp
                workbook.CalculateFormula(calcOptions);

                // Retrieve the audit XML
                XDocument auditXml = monitor.GetAuditXml();

                // Save the XML report to a file
                string outputXmlPath = "FormulaAuditTrail.xml";
                using (FileStream fs = new FileStream(outputXmlPath, FileMode.Create, FileAccess.Write))
                {
                    auditXml.Save(fs);
                }

                Console.WriteLine($"Formula audit trail exported to '{outputXmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
