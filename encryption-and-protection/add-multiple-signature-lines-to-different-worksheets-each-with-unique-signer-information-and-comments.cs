using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddMultipleSignatureLines
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and rename it
        Worksheet financeSheet = workbook.Worksheets[0];
        financeSheet.Name = "Finance";

        // Add two more worksheets
        Worksheet hrSheet = workbook.Worksheets.Add("HR");
        Worksheet itSheet = workbook.Worksheets.Add("IT");

        // -------------------- Finance sheet signature line --------------------
        SignatureLine financeSignature = new SignatureLine();
        financeSignature.Signer = "Alice Johnson";
        financeSignature.Title = "Finance Manager";
        financeSignature.Email = "alice.johnson@company.com";
        financeSignature.Instructions = "Please sign to approve the financial report.";
        financeSignature.AllowComments = true;
        financeSignature.ShowSignedDate = true;
        financeSignature.IsLine = true;

        // Add the signature line at row 2, column 2 (zero‑based indices)
        financeSheet.Shapes.AddSignatureLine(1, 1, financeSignature);

        // -------------------- HR sheet signature line --------------------
        SignatureLine hrSignature = new SignatureLine();
        hrSignature.Signer = "Bob Smith";
        hrSignature.Title = "HR Director";
        hrSignature.Email = "bob.smith@company.com";
        hrSignature.Instructions = "Sign to confirm HR policies.";
        hrSignature.AllowComments = false;
        hrSignature.ShowSignedDate = true;
        hrSignature.IsLine = true;

        // Add the signature line at row 4, column 3
        hrSheet.Shapes.AddSignatureLine(3, 2, hrSignature);

        // -------------------- IT sheet signature line --------------------
        SignatureLine itSignature = new SignatureLine();
        itSignature.Signer = "Carol Lee";
        itSignature.Title = "IT Lead";
        itSignature.Email = "carol.lee@company.com";
        itSignature.Instructions = "Approve the IT infrastructure changes.";
        itSignature.AllowComments = true;
        itSignature.ShowSignedDate = false;
        itSignature.IsLine = true;

        // Add the signature line at row 6, column 1
        itSheet.Shapes.AddSignatureLine(5, 0, itSignature);

        // Save the workbook with all signature lines
        workbook.Save("MultipleSignatureLines.xlsx");
    }
}