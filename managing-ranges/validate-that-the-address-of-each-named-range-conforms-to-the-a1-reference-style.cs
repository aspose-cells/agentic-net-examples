using System;
using Aspose.Cells;

namespace NamedRangeA1Validation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create rule

            // Add sample data and named ranges for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // A1 style named range (valid)
            int idx1 = workbook.Worksheets.Names.Add("ValidRange");
            Name validName = workbook.Worksheets.Names[idx1];
            validName.RefersTo = "=Sheet1!$A$1:$A$3";

            // R1C1 style named range (invalid for A1)
            int idx2 = workbook.Worksheets.Names.Add("InvalidRange");
            Name invalidName = workbook.Worksheets.Names[idx2];
            // Set using R1C1 reference
            invalidName.R1C1RefersTo = "'Sheet1'!R1C1:R3C1";

            // Iterate through all defined names and verify A1 reference style
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Original reference string as stored
                string storedRef = name.RefersTo;

                // Get the reference formatted explicitly in A1 style
                string a1Ref = name.GetRefersTo(false, false); // get A1, global/local not relevant here

                // Compare the two strings (case‑insensitive)
                bool isA1 = string.Equals(storedRef, a1Ref, StringComparison.OrdinalIgnoreCase);

                Console.WriteLine($"Name: {name.Text}");
                Console.WriteLine($"Stored RefersTo: {storedRef}");
                Console.WriteLine($"A1 RefersTo   : {a1Ref}");
                Console.WriteLine($"Conforms to A1 style: {(isA1 ? "Yes" : "No")}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook (save rule)
            workbook.Save("NamedRangeA1Validation_Output.xlsx");
        }
    }
}