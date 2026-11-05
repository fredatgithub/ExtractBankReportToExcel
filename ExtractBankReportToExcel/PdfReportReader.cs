using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace ExtractBankReportToExcel;

internal static partial class PdfReportReader
{
  private static readonly CultureInfo FrenchCulture = CultureInfo.GetCultureInfo("fr-FR");

  internal static IReadOnlyList<BankOperation> Read(string pdfPath)
  {
    using var document = PdfDocument.Open(pdfPath);
    var operations = new List<BankOperation>();
    var (monthlyPaymentAmount, outstandingBalance, installmentMonthlyPaymentAmount) = ReadStatementAmounts(document);

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

          if (TryReadOperation(
            line,
            Path.GetFileName(pdfPath),
            monthlyPaymentAmount,
            outstandingBalance,
            installmentMonthlyPaymentAmount,
            out var operation))
          {
            operations.Add(operation);
          }
        }
      }
    }

    return operations;
  }

  private static (decimal? MonthlyPaymentAmount, decimal? OutstandingBalance, decimal? InstallmentMonthlyPaymentAmount) ReadStatementAmounts(PdfDocument document)
  {
    decimal? monthlyPaymentAmount = null;
    decimal? outstandingBalance = null;
    decimal? installmentMonthlyPaymentAmount = null;

    foreach (var page in document.GetPages())
    {
      var lines = page.GetWords()
        .GroupBy(word => Math.Round(word.BoundingBox.BottomLeft.Y))
        .OrderByDescending(line => line.Key)
        .ToList();

      for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
      {
        var lineText = GetNormalizedLineText(lines[lineIndex]);

        if (monthlyPaymentAmount is null
          && lineText.Contains("MONTANT", StringComparison.Ordinal)
          && lineText.Contains("PROCHAIN", StringComparison.Ordinal)
          && lineText.Contains("PRELEVEMENT", StringComparison.Ordinal)
          && TryReadLabeledAmount(lines[lineIndex], out var paymentAmount))
        {
          monthlyPaymentAmount = paymentAmount;
        }

        if (!IsInstallmentCreditSection(lineText) || (outstandingBalance is not null && installmentMonthlyPaymentAmount is not null))
        {
          continue;
        }

        for (var sectionLineIndex = lineIndex + 1; sectionLineIndex < lines.Count; sectionLineIndex++)
        {
          var sectionLineText = GetNormalizedLineText(lines[sectionLineIndex]);
          if (IsCreditSection(sectionLineText))
          {
            break;
          }

          if (outstandingBalance is null
            && IsOutstandingBalanceLabel(sectionLineText)
            && TryReadLabeledAmount(lines[sectionLineIndex], out var balance))
          {
            outstandingBalance = balance;
          }

          if (installmentMonthlyPaymentAmount is null
            && IsInstallmentMonthlyPaymentLabel(sectionLineText)
            && TryReadLabeledAmount(lines[sectionLineIndex], out var installmentMonthlyPayment))
          {
            installmentMonthlyPaymentAmount = installmentMonthlyPayment;
          }

          if (outstandingBalance is not null && installmentMonthlyPaymentAmount is not null)
          {
            break;
          }
        }
      }
    }

    return (monthlyPaymentAmount, outstandingBalance, installmentMonthlyPaymentAmount);
  }

  private static bool IsInstallmentCreditSection(string normalizedText)
  {
    return normalizedText.Contains("MES OPERATIONS", StringComparison.Ordinal)
      && normalizedText.Contains("EN PLUSIEURS FOIS", StringComparison.Ordinal)
      && normalizedText.Contains("A CREDIT", StringComparison.Ordinal);
  }

  private static bool IsCreditSection(string normalizedText)
  {
    return normalizedText.Contains("MES OPERATIONS", StringComparison.Ordinal)
      && normalizedText.Contains("A CREDIT", StringComparison.Ordinal);
  }

  private static bool IsOutstandingBalanceLabel(string normalizedText)
  {
    return normalizedText.Contains("MONTANT", StringComparison.Ordinal)
      && normalizedText.Contains("RESTANT", StringComparison.Ordinal)
      && normalizedText.Contains("DU", StringComparison.Ordinal);
  }

  private static bool IsInstallmentMonthlyPaymentLabel(string normalizedText)
  {
    return normalizedText.Contains("MA MENSUALITE", StringComparison.Ordinal);
  }

  private static string GetNormalizedLineText(IEnumerable<Word> words)
  {
    return string.Join(" ", words.Select(word => Normalize(word.Text)));
  }

  private static bool TryReadLabeledAmount(IEnumerable<Word> words, out decimal amount)
  {
    var amountText = string.Concat(words
        .Where(word => word.BoundingBox.BottomLeft.X >= 500)
        .OrderBy(word => word.BoundingBox.BottomLeft.X)
        .Select(word => word.Text)
        .Where(token => Regex.IsMatch(token, @"^[\d.,\-−€]+$") && token.Any(char.IsDigit))
        .Select(token => token.Replace("€", string.Empty, StringComparison.Ordinal)));

    return TryParseAmount(amountText, out amount);
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
    decimal? monthlyPaymentAmount,
    decimal? outstandingBalance,
    decimal? installmentMonthlyPaymentAmount,
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
      MonthlyPaymentAmount = monthlyPaymentAmount,
      OutstandingBalance = outstandingBalance,
      InstallmentMonthlyPaymentAmount = installmentMonthlyPaymentAmount,
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
