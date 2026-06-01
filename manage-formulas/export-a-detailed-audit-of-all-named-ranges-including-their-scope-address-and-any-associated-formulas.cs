using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class NamedRangeAudit
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data and named ranges (for demonstration)
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            // Global (workbook‑level) named range
            int globalIdx = workbook.Worksheets.Names.Add("GlobalRange");
            Name globalName = workbook.Worksheets.Names[globalIdx];
            globalName.RefersTo = "=Sheet1!$A$1:$A$3";

            // Worksheet‑scoped named range
            int localIdx = workbook.Worksheets.Names.Add("Sheet1!LocalRange");
            Name localName = workbook.Worksheets.Names[localIdx];
            localName.RefersTo = "=Sheet1!$B$1:$B$3";
            localName.SheetIndex = 0; // 0 = global, otherwise 1‑based sheet index

            // Add a second sheet with its own named range
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            sheet2.Cells["C1"].PutValue(100);
            sheet2.Cells["C2"].PutValue(200);
            int dataIdx = workbook.Worksheets.Names.Add("DataRange");
            Name dataName = workbook.Worksheets.Names[dataIdx];
            dataName.RefersTo = "=Data!$C$1:$C$2";

            // -------------------------------------------------
            // Create a worksheet to store the audit report
            // -------------------------------------------------
            Worksheet auditSheet = workbook.Worksheets.Add("Audit");
            int auditRow = 0;
            auditSheet.Cells[auditRow, 0].PutValue("Name");
            auditSheet.Cells[auditRow, 1].PutValue("Scope");
            auditSheet.Cells[auditRow, 2].PutValue("Address");
            auditSheet.Cells[auditRow, 3].PutValue("RefersTo");
            auditRow++;

            // -------------------------------------------------
            // Enumerate all defined names and collect details
            // -------------------------------------------------
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Determine the scope (global or specific worksheet)
                string scope = name.SheetIndex == 0
                    ? "Workbook (Global)"
                    : $"Worksheet: {workbook.Worksheets[name.SheetIndex - 1].Name}";

                // Try to obtain the actual range address (if the name refers to a range)
                string address = string.Empty;
                try
                {
                    AsposeRange rng = name.GetRange(); // uses Name.GetRange()
                    if (rng != null)
                    {
                        address = rng.Address;
                    }
                }
                catch
                {
                    // If GetRange throws, the name does not refer to a simple range
                    address = "N/A";
                }

                // The formula that defines the name
                string refersTo = name.RefersTo;

                // Output to console
                Console.WriteLine($"Name: {name.Text}, Scope: {scope}, Address: {address}, RefersTo: {refersTo}");

                // Write the same information into the audit worksheet
                auditSheet.Cells[auditRow, 0].PutValue(name.Text);
                auditSheet.Cells[auditRow, 1].PutValue(scope);
                auditSheet.Cells[auditRow, 2].PutValue(address);
                auditSheet.Cells[auditRow, 3].PutValue(refersTo);
                auditRow++;
            }

            // -------------------------------------------------
            // Save the workbook containing the audit sheet
            // -------------------------------------------------
            string outputPath = "NamedRangeAudit.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Audit workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}