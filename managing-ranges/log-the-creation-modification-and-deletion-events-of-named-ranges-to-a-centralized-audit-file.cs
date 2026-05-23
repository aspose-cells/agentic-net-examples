using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeAuditDemo
{
    class Program
    {
        // Path to the centralized audit file
        private const string AuditFilePath = "NamedRangeAudit.log";

        // Append a message with timestamp to the audit file
        static void Log(string message)
        {
            try
            {
                string entry = $"{DateTime.UtcNow:O} - {message}";
                File.AppendAllLines(AuditFilePath, new[] { entry });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging error: {ex.Message}");
            }
        }

        static void Main()
        {
            try
            {
                // Ensure the audit file is clean for this run
                if (File.Exists(AuditFilePath))
                {
                    File.Delete(AuditFilePath);
                }

                // -------------------- Create Workbook --------------------
                Workbook workbook = new Workbook(); // create new workbook
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data
                sheet.Cells["A1"].PutValue("Apple");
                sheet.Cells["A2"].PutValue("Banana");
                sheet.Cells["A3"].PutValue("Cherry");

                // -------------------- Create Named Range --------------------
                // Use the workbook's global Names collection
                NameCollection names = workbook.Worksheets.Names;

                // Define a new name "Fruits"
                int nameIndex = names.Add("Fruits");
                Name fruitName = names[nameIndex];
                // Use the actual worksheet name (default is "Sheet1")
                string sheetName = sheet.Name;
                fruitName.RefersTo = $"={sheetName}!$A$1:$A$3";

                Log($"Created named range '{fruitName.Text}' referring to {fruitName.RefersTo}");

                // -------------------- Modify Named Range --------------------
                // Change the reference to include an extra row
                fruitName.RefersTo = $"={sheetName}!$A$1:$A$4";
                Log($"Modified named range '{fruitName.Text}' new reference {fruitName.RefersTo}");

                // Add the extra data row
                sheet.Cells["A4"].PutValue("Date");

                // -------------------- Delete Named Range --------------------
                // Remove the named range from the collection
                names.Remove(fruitName.Text);
                Log($"Deleted named range '{fruitName.Text}'");

                // -------------------- Save Workbook --------------------
                string outputPath = "NamedRangeAuditDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Log($"Error saving workbook: {ex.Message}");
                }

                // -------------------- Output Audit Log --------------------
                Console.WriteLine("Audit log entries:");
                if (File.Exists(AuditFilePath))
                {
                    Console.WriteLine(File.ReadAllText(AuditFilePath));
                }
                else
                {
                    Console.WriteLine("No audit log found.");
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}