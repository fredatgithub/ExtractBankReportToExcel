namespace ExtractBankReportToExcel;

public sealed class BankOperation
{
  public DateTime? OperationDate { get; init; }

  public DateTime? ValueDate { get; init; }

  public string Description { get; init; } = string.Empty;

  public decimal? Debit { get; init; }

  public decimal? Credit { get; init; }

  public decimal? MonthlyPaymentAmount { get; init; }

  public decimal? OutstandingBalance { get; init; }

  public decimal? InstallmentMonthlyPaymentAmount { get; init; }

  public int? RemainingInstallmentCount
  {
    get
    {
      if (OutstandingBalance is not decimal balance
        || InstallmentMonthlyPaymentAmount is not decimal monthlyPayment
        || balance < 0
        || monthlyPayment <= 0)
      {
        return null;
      }

      var months = decimal.Ceiling(balance / monthlyPayment);
      return months <= int.MaxValue ? (int)months : null;
    }
  }

  public string? RemainingYearsAndMonths
  {
    get
    {
      if (RemainingInstallmentCount is not int months)
      {
        return null;
      }

      var years = months / 12;
      var remainingMonths = months % 12;
      if (years == 0)
      {
        return $"{months} mois";
      }

      var yearLabel = years == 1 ? "1 an" : $"{years} ans";
      return remainingMonths == 0 ? yearLabel : $"{yearLabel} et {remainingMonths} mois";
    }
  }

  public string SourceFile { get; init; } = string.Empty;
}
