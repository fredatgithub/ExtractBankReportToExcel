using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace ExtractBankReportToExcel;

internal static partial class PdfReportReader
{
  private static readonly CultureInfo FrenchCulture = CultureInfo.GetCultureInfo("fr-FR");

  internal static IReadOnlyList<BankOperation> Read(string pdfPath)
  {
    using var document = PdfDocument.Open(pdfPath);
    var operations = new List<BankOperation>();

    foreach (var page in document.GetPages())
    {
      var lines = page.GetWords()
        .GroupBy(word => Math.Round(word.BoundingBox.BottomLeft.Y))
        .OrderByDescending(line => line.Key)
        .ToList();

      for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
      {
        if (!IsOperationsHeader(lines[lineIndex]))
        {
          continue;
        }

        for (var operationIndex = lineIndex + 1; operationIndex < lines.Count; operationIndex++)
        {
          var line = lines[operationIndex];
          if (IsOperationsHeader(line))
          {
            break;
          }

          if (TryReadOperation(line, Path.GetFileName(pdfPath), out var operation))
          {
            operations.Add(operation);
          }
        }
      }
    }

    return operations;
  }

  private static bool IsOperationsHeader(IEnumerable<Word> words)
  {
    var normalizedText = string.Join(" ", words.Select(word => Normalize(word.Text)));
    return normalizedText.Contains("DATE", StringComparison.Ordinal)
      && normalizedText.Contains("OPERATION", StringComparison.Ordinal)
      && normalizedText.Contains("DEBIT", StringComparison.Ordinal)
      && normalizedText.Contains("CREDIT", StringComparison.Ordinal);
  }

  private static bool TryReadOperation(
    IGrouping<double, Word> line,
    string sourceFile,
    out BankOperation operation)
  {
    var words = line.OrderBy(word => word.BoundingBox.BottomLeft.X).ToList();
    var dates = words
      .Where(word => word.BoundingBox.BottomLeft.X < 160)
      .Select(word => (Word: word, Value: ParseDate(word.Text)))
      .Where(item => item.Value.HasValue)
      .Select(item => item.Value!.Value)
      .ToList();

    var description = string.Join(
      " ",
      words
        .Where(word => word.BoundingBox.BottomLeft.X is >= 160 and < 430)
        .Where(word => !TryParseAmount(word.Text, out _))
        .Select(word => word.Text));

    var amounts = words
      .Where(word => word.BoundingBox.BottomLeft.X >= 430)
      .Select(word => (X: word.BoundingBox.BottomLeft.X, Parsed: TryParseAmount(word.Text, out var value), Value: value))
      .Where(item => item.Parsed)
      .ToList();

    var debit = amounts.Where(amount => amount.X < 500).Select(amount => (decimal?)amount.Value).FirstOrDefault();
    var credit = amounts.Where(amount => amount.X >= 500).Select(amount => (decimal?)amount.Value).FirstOrDefault();

    if (dates.Count == 0 || string.IsNullOrWhiteSpace(description) || (debit is null && credit is null))
    {
      operation = null!;
      return false;
    }

    operation = new BankOperation
    {
      OperationDate = dates[0],
      ValueDate = dates.Count > 1 ? dates[1] : null,
      Description = description,
      Debit = debit,
      Credit = credit,
      SourceFile = sourceFile
    };
    return true;
  }

  private static DateTime? ParseDate(string value)
  {
    string[] formats =
    [
      "d/M/yyyy", "dd/MM/yyyy", "d/M/yy", "dd/MM/yy",
      "d-M-yyyy", "dd-MM-yyyy", "d-M-yy", "dd-MM-yy"
    ];

    return DateTime.TryParseExact(
      value.Trim(),
      formats,
      FrenchCulture,
      DateTimeStyles.None,
      out var date)
      ? date
      : null;
  }

  private static bool TryParseAmount(string value, out decimal amount)
  {
    var normalized = value
      .Replace("€", string.Empty, StringComparison.Ordinal)
      .Replace("\u00A0", string.Empty, StringComparison.Ordinal)
      .Replace(" ", string.Empty, StringComparison.Ordinal)
      .Trim();

    return decimal.TryParse(
      normalized,
      NumberStyles.Number | NumberStyles.AllowLeadingSign,
      FrenchCulture,
      out amount);
  }

  private static string Normalize(string value)
  {
    var decomposed = value.Normalize(NormalizationForm.FormD);
    var builder = new StringBuilder(decomposed.Length);

    foreach (var character in decomposed)
    {
      if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
      {
        builder.Append(char.ToUpperInvariant(character));
      }
    }

    return builder.ToString();
  }
}
