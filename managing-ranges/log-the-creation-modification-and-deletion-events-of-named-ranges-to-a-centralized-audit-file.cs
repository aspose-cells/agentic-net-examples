using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeAuditDemo
{
    // Simple logger that appends audit entries to a text file
    public static class AuditLogger
    {
        private static readonly string LogFilePath = "NamedRangeAudit.log";

        public static void Log(string message)
        {
            string entry = $"{DateTime.UtcNow:O} - {message}";
            File.AppendAllLines(LogFilePath, new[] { entry });
        }
    }

    public static class NamedRangeHelper
    {
        // Creates a named range and logs the creation event
        public static void AddNamedRange(Workbook workbook, string name, string refersTo)
        {
            // Add the name to the collection (global scope)
            int index = workbook.Worksheets.Names.Add(name);
            Name definedName = workbook.Worksheets.Names[index];
            definedName.RefersTo = refersTo;

            AuditLogger.Log($"Created named range '{name}' with RefersTo = \"{refersTo}\"");
        }

        // Modifies an existing named range and logs the modification event
        public static void ModifyNamedRange(Workbook workbook, string name, string newRefersTo)
        {
            Name definedName = workbook.Worksheets.Names[name];
            if (definedName == null)
            {
                AuditLogger.Log($"Attempted to modify non‑existent named range '{name}'");
                return;
            }

            string oldRefersTo = definedName.RefersTo;
            definedName.RefersTo = newRefersTo;

            AuditLogger.Log($"Modified named range '{name}': RefersTo changed from \"{oldRefersTo}\" to \"{newRefersTo}\"");
        }

        // Removes a named range and logs the deletion event
        public static void RemoveNamedRange(Workbook workbook, string name)
        {
            Name definedName = workbook.Worksheets.Names[name];
            if (definedName == null)
            {
                AuditLogger.Log($"Attempted to delete non‑existent named range '{name}'");
                return;
            }

            workbook.Worksheets.Names.Remove(name);
            AuditLogger.Log($"Deleted named range '{name}'");
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            Workbook wb = new Workbook();                     // Create a new workbook
            Worksheet ws = wb.Worksheets[0];                  // Access first worksheet

            // Populate some data to be referenced by named ranges
            ws.Cells["A1"].PutValue("Item");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["A3"].PutValue("Banana");

            // Create a named range and log the creation
            NamedRangeHelper.AddNamedRange(wb, "FruitList", "=Sheet1!$A$2:$A$3");

            // ---------- Modify ----------
            // Change the range to include an extra cell
            NamedRangeHelper.ModifyNamedRange(wb, "FruitList", "=Sheet1!$A$2:$A$4");

            // Add another cell so the new reference is valid
            ws.Cells["A4"].PutValue("Cherry");

            // ---------- Delete ----------
            // Remove the named range and log the deletion
            NamedRangeHelper.RemoveNamedRange(wb, "FruitList");

            // ---------- Save ----------
            wb.Save("NamedRangeAuditDemo.xlsx");              // Save the workbook

            // Optional: Load the workbook back to demonstrate load lifecycle
            Workbook loadedWb = new Workbook("NamedRangeAuditDemo.xlsx");
            // No further actions; the audit log already contains the events
        }
    }
}