namespace ExtractBankReportToExcel;

public sealed class BankOperation
{
  public DateTime? OperationDate { get; init; }

  public DateTime? ValueDate { get; init; }

  public string Description { get; init; } = string.Empty;

  public decimal? Debit { get; init; }

  public decimal? Credit { get; init; }

  public string SourceFile { get; init; } = string.Empty;
}
