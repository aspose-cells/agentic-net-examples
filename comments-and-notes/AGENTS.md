---
name: Aspose.Cells Comments and Notes Agent
category: comments-and-notes
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Add, read, update, format, copy, and remove Excel comments and threaded comments in C#
primary_apis: [Comment, CommentCollection, ThreadedComment, ThreadedCommentCollection, ThreadedCommentAuthor]
related_categories: [../cells-data/, ../format-cells/, ../working-with-shapes/]
---

# Comments and Notes Agent Instructions

## Mission and scope

Create correct, focused examples for legacy cell comments/notes and modern threaded comments with Aspose.Cells for .NET. Follow [`../AGENTS.md`](../AGENTS.md) first.

In scope: adding, reading, editing, copying, formatting, enumerating, and removing comments; authorship; threaded conversations; comment shapes; text direction; visibility; and comment audits.

Keep general shape formatting in `working-with-shapes` and ordinary cell text in `cells-data` unless comments are the primary intent.

## Model the annotation type explicitly

| Intent | APIs |
| --- | --- |
| Legacy comment/note | `Worksheet.Comments`, `CommentCollection`, `Comment` |
| Threaded comment | `Comment.ThreadedComments`, `CommentCollection.AddThreadedComment`, `ThreadedCommentCollection` |
| Threaded author | `Workbook.Worksheets.ThreadedCommentAuthors` and `ThreadedCommentAuthor` |
| Comment text | `Comment.Note` and version-supported HTML/text properties |
| Appearance | `Comment.CommentShape` and verified shape/font APIs |
| Removal | Collection removal APIs verified for the annotation type |

Do not call a legacy note a threaded comment or imply that their authorship, timestamps, replies, or storage models are interchangeable.

## Hard rules

- Attach comments to valid cells and verify the returned collection index/object.
- Use synthetic author names and text; never place personal data in generated examples.
- Preserve author, text, formatting, and thread order only when the chosen API supports them.
- Treat comment text as untrusted input when exporting to HTML, logs, or other systems.
- Do not infer creation times, reply APIs, or removal methods from filenames; verify the installed API.
- Save threaded-comment scenarios to a format that preserves them, normally XLSX.

## Canonical legacy-comment pattern

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

int index = worksheet.Comments.Add("A1");
Comment comment = worksheet.Comments[index];
comment.Note = "Review the quarterly total.";

if (comment.Note != "Review the quarterly total.")
{
    throw new InvalidOperationException("Comment text was not retained.");
}

workbook.Save("comment-result.xlsx");
```

## Example contract and safety

Each example must identify whether it uses a legacy or threaded comment, demonstrate one primary operation, use explicit types, verify cell/author/text/count, and reopen output when persistence is the subject.

Use metadata fields for title, intent, annotation type, primary API, target cell, output, and expected result. Prefer filenames such as `add-threaded-comment-to-excel-cell.cs`.

Avoid external links, real email addresses, confidential review notes, and HTML injection. Do not execute or interpret comment text as code.

## Discoverability and validation

Target one question such as "add a comment to Excel in C#," "read threaded comments," or "format an Excel note." The opening comment must directly identify the annotation type and expected cell.

Verify exact author, collection, shape, and removal APIs against the installed package. Compile, run, reopen, and confirm comment count, text, author, and target cell. Reject examples that conflate comments with cell values or silently lose threaded metadata.

## Related knowledge

- [Category overview](README.md)
- [Cell data](../cells-data/)
- [Cell formatting](../format-cells/)
- [Shapes](../working-with-shapes/)
- [Official comments documentation](https://docs.aspose.com/cells/net/comments-and-notes/)

## Definition of done

The example is done when the annotation type and API are unambiguous, synthetic content is safe, the saved workbook preserves the verified text/author/cell relationship, and the intent is immediately retrievable.

