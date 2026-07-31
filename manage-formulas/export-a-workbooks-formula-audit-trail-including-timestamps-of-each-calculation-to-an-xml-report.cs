// Title: Export Formula Audit Trail with Timestamps to XML using Aspose.Cells for .NET
// Description: Loads an Excel workbook, switches to manual calculation mode, and uses a custom AbstractCalculationMonitor to capture each formula cell and the UTC time it was evaluated. After calculation the data is written to a well‑formed XML file that lists cell address, formula text, and ISO‑8601 timestamp.
// Keywords: Aspose.Cells | .NET | C# | formula audit | calculation timestamps | XML report | AbstractCalculationMonitor | manual calculation mode | Excel audit log | global compliance tracking
// Common Searches: Aspose.Cells export formula audit to XML | log formula calculation time C# | custom calculation monitor example | generate XML audit report for Excel formulas | record formula evaluation timestamps Aspose.Cells
// Developer Intent: Create an XML file that records every evaluated formula cell together with the exact UTC timestamp of its calculation.
// Use Cases: Compliance auditing: prove when each formula was calculated during a batch run. | Performance monitoring: identify formulas that trigger unexpectedly frequent recalculations. | Integration with logging pipelines: feed the XML audit into downstream monitoring or alerting systems. | Version control of spreadsheet logic: keep a timestamped snapshot of all formulas after a manual calculation.
// AI Prompts: Show how to extend AuditCalculationMonitor to also capture the worksheet name in the XML output. | Provide code that reads FormulaAuditReport.xml and aggregates the number of formulas per worksheet. | Explain how to convert the UTC timestamps to a specific local time zone while preserving ISO‑8601 format.

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsAudit
{
    // Custom monitor to record formula evaluation timestamps
    // Loads an Excel workbook, switches to manual calculation mode, and uses a custom AbstractCalculationMonitor to capture each formula cell and the UTC time it was evaluated. After calculation the data is written to a well‑formed XML file that lists cell address, formula text, and ISO‑8601 timestamp.
    class AuditCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly Dictionary<string, (string Formula, DateTime Timestamp)> _audit;
        private readonly WorksheetCollection _worksheets;

        public AuditCalculationMonitor(Dictionary<string, (string Formula, DateTime Timestamp)> audit, WorksheetCollection worksheets)
        {
            _audit = audit;
            _worksheets = worksheets;
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Worksheet sheet = _worksheets[sheetIndex];
            Cell cell = sheet.Cells[rowIndex, columnIndex];

            // Record only cells that contain a formula
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                string address = cell.Name; // e.g., "A1"
                string formula = cell.Formula;
                DateTime timestamp = DateTime.UtcNow; // UTC for consistency

                _audit[address] = (formula, timestamp);
            }
        }

        public override bool OnCircular(System.Collections.IEnumerator circularCellsData) => base.OnCircular(circularCellsData);
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex) { }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string workbookPath = "InputWorkbook.xlsx";

                // Verify the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Use manual calculation mode to control when formulas are evaluated
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Dictionary to hold audit information: cell address -> (formula, timestamp)
                var auditInfo = new Dictionary<string, (string Formula, DateTime Timestamp)>();

                // Set up calculation options with the custom monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = new AuditCalculationMonitor(auditInfo, workbook.Worksheets)
                };

                // Perform the calculation; this triggers the monitor for each formula cell
                workbook.CalculateFormula(calcOptions);

                // Generate the XML audit report
                string xmlReportPath = "FormulaAuditReport.xml";
                using (XmlWriter writer = XmlWriter.Create(xmlReportPath, new XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("FormulaAudit");

                    foreach (var entry in auditInfo)
                    {
                        writer.WriteStartElement("Cell");
                        writer.WriteAttributeString("Address", entry.Key);
                        writer.WriteAttributeString("Formula", entry.Value.Formula);
                        writer.WriteAttributeString("Timestamp", entry.Value.Timestamp.ToString("o")); // ISO 8601
                        writer.WriteEndElement(); // Cell
                    }

                    writer.WriteEndElement(); // FormulaAudit
                    writer.WriteEndDocument();
                }

                Console.WriteLine($"Formula audit trail exported to '{xmlReportPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
