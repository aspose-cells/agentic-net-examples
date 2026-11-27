// Title: C# program to validate unique email addresses in column AK and highlight duplicates using Aspose.Cells
// Common Searches: C# Aspose.Cells how to find duplicate email addresses in column AK of an XLSX file | mark duplicate emails with red background and comment using Aspose.Cells .NET | validate email format and uniqueness in Excel column AK with Aspose.Cells | Aspose.Cells example for detecting and flagging duplicate email cells in C# | using HashSet to enforce unique email values in an Excel worksheet via Aspose.Cells
// Tags: detect duplicate emails Aspose.Cells | highlight duplicate cells red background Aspose.Cells | add cell comment for validation error Aspose.Cells | email regex validation Aspose.Cells C# | column AK email uniqueness Aspose.Cells | hashset case-insensitive email tracking Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace UniqueEmailValidator
{
    // The sample loads an XLSX workbook with Aspose.Cells, identifies column AK, iterates through each populated row, checks the cell value against a simple email regex, and uses a case‑insensitive HashSet to detect duplicates. Cells with invalid format or duplicate addresses are marked with a red background style and a comment describing the error. The workbook is then saved to a new file, providing a quick way to enforce unique email addresses in a specific Excel column.
    class Program
    {
        // Simple e‑mail pattern – not exhaustive but sufficient for demonstration
        private static readonly Regex EmailRegex = new Regex(
            @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // Input / Output handling
            // -----------------------------------------------------------------
            string inputPath;
            string outputPath;

            if (args.Length >= 1)
            {
                inputPath = args[0];
                outputPath = args.Length >= 2
                    ? args[1]
                    : Path.Combine(
                        Path.GetDirectoryName(inputPath) ?? Environment.CurrentDirectory,
                        Path.GetFileNameWithoutExtension(inputPath) + "_validated.xlsx");
            }
            else
            {
                Console.WriteLine("Usage: UniqueEmailValidator <input.xlsx> [output.xlsx]");
                return;
            }

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // -----------------------------------------------------------------
                // Load workbook
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0]; // Assuming validation on the first sheet
                Cells cells = worksheet.Cells;

                // -----------------------------------------------------------------
                // Determine column index for "AK"
                // -----------------------------------------------------------------
                int emailColumnIndex = CellsHelper.ColumnNameToIndex("AK"); // 0‑based index

                // -----------------------------------------------------------------
                // Prepare structures for duplicate detection
                // -----------------------------------------------------------------
                var seenEmails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var duplicateCells = new List<Cell>();

                // -----------------------------------------------------------------
                // Scan the column
                // -----------------------------------------------------------------
                int firstDataRow = cells.MinRow; // Usually 0
                int lastDataRow = cells.MaxDataRow; // Last row that contains data

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    Cell cell = cells[row, emailColumnIndex];
                    if (cell == null || cell.Type == CellValueType.IsNull)
                        continue; // Empty cell – nothing to validate

                    string email = cell.StringValue?.Trim();
                    if (string.IsNullOrEmpty(email))
                        continue; // Blank cell – ignore

                    // Optional: basic e‑mail format check
                    if (!EmailRegex.IsMatch(email))
                    {
                        MarkCellAsInvalid(cell, "Invalid e‑mail format");
                        continue;
                    }

                    // Duplicate detection
                    if (!seenEmails.Add(email))
                    {
                        // Email already existed – mark as duplicate
                        duplicateCells.Add(cell);
                    }
                }

                // -----------------------------------------------------------------
                // Apply visual error indication to duplicates
                // -----------------------------------------------------------------
                foreach (Cell dupCell in duplicateCells)
                {
                    MarkCellAsInvalid(dupCell, "Duplicate e‑mail address");
                }

                // -----------------------------------------------------------------
                // Save the workbook
                // -----------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Validation complete. Result saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }

        /// <param name="cell">The cell to mark.</param>
        /// <param name="message">Error message to display.</param>
        private static void MarkCellAsInvalid(Cell cell, string message)
        {
            try
            {
                // Apply red background style
                Style style = cell.GetStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);

                // Add a comment to the cell's worksheet
                Worksheet sheet = cell.Worksheet;
                string cellName = CellsHelper.CellIndexToName(cell.Row, cell.Column);
                int commentIndex = sheet.Comments.Add(cellName);
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = message;
                comment.Author = "Validator";
                // No explicit positioning needed; the comment is attached to the cell.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to mark cell at ({cell.Row}, {cell.Column}): {ex.Message}");
            }
        }
    }
}
