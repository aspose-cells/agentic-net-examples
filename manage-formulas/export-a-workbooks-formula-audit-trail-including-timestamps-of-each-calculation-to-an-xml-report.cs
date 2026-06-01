using System;
using System.Collections.Generic;
using System.Xml;
using Aspose.Cells;

// Record for each cell calculation
class CalculationRecord
{
    public string SheetName { get; set; }
    public string CellAddress { get; set; }
    public DateTime Timestamp { get; set; }
    public string Value { get; set; }
}

// Custom calculation monitor to capture audit information
class AuditCalculationMonitor : AbstractCalculationMonitor
{
    private readonly Workbook _workbook;
    private readonly List<CalculationRecord> _records;

    public AuditCalculationMonitor(Workbook workbook, List<CalculationRecord> records)
    {
        _workbook = workbook;
        _records = records;
    }

    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        var sheet = _workbook.Worksheets[sheetIndex];
        var cell = sheet.Cells[rowIndex, columnIndex];
        var record = new CalculationRecord
        {
            SheetName = sheet.Name,
            CellAddress = cell.Name,
            Timestamp = DateTime.Now,
            Value = cell.Value?.ToString() ?? string.Empty
        };
        _records.Add(record);
    }
}

class FormulaAuditExport
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Set calculation mode to Manual to control when formulas are evaluated
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prepare a list to hold audit records
        List<CalculationRecord> auditRecords = new List<CalculationRecord>();

        // Configure calculation options with the custom monitor
        CalculationOptions calcOptions = new CalculationOptions();
        calcOptions.CalculationMonitor = new AuditCalculationMonitor(workbook, auditRecords);

        // Perform formula calculation while the monitor records each cell's evaluation
        workbook.CalculateFormula(calcOptions);

        // Export the audit trail to an XML file
        using (XmlWriter writer = XmlWriter.Create("FormulaAuditTrail.xml", new XmlWriterSettings { Indent = true }))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("FormulaAuditTrail");

            foreach (var rec in auditRecords)
            {
                writer.WriteStartElement("Calculation");
                writer.WriteElementString("Sheet", rec.SheetName);
                writer.WriteElementString("Cell", rec.CellAddress);
                writer.WriteElementString("Timestamp", rec.Timestamp.ToString("o")); // ISO 8601 format
                writer.WriteElementString("Value", rec.Value);
                writer.WriteEndElement(); // Calculation
            }

            writer.WriteEndElement(); // FormulaAuditTrail
            writer.WriteEndDocument();
        }

        Console.WriteLine("Formula audit trail exported to 'FormulaAuditTrail.xml'.");
    }
}