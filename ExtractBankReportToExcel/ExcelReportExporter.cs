using ClosedXML.Excel;

namespace ExtractBankReportToExcel;

internal static class ExcelReportExporter
{
  internal static void Export(string filePath, IReadOnlyList<BankOperation> operations)
  {
    using var workbook = new XLWorkbook();
    var worksheet = workbook.Worksheets.Add("Opérations");
    var headers = new[]
    {
      "Date opération", "Date valeur", "Libellé / opération", "Débit", "Crédit",
      "Montant prélèvement", "Montant restant dû", "Fichier PDF",
      "Nombre de mensualités restant dues", "Nombre d'années restantes"
    };

    for (var column = 0; column < headers.Length; column++)
    {
      worksheet.Cell(1, column + 1).Value = headers[column];
    }

    for (var index = 0; index < operations.Count; index++)
    {
      var operation = operations[index];
      var row = index + 2;
      if (operation.OperationDate is DateTime operationDate)
      {
        worksheet.Cell(row, 1).Value = operationDate;
      }

      if (operation.ValueDate is DateTime valueDate)
      {
        worksheet.Cell(row, 2).Value = valueDate;
      }

      worksheet.Cell(row, 3).Value = operation.Description;
      if (operation.Debit is decimal debit)
      {
        worksheet.Cell(row, 4).Value = debit;
      }

      if (operation.Credit is decimal credit)
      {
        worksheet.Cell(row, 5).Value = credit;
      }

      if (operation.MonthlyPaymentAmount is decimal monthlyPaymentAmount)
      {
        worksheet.Cell(row, 6).Value = monthlyPaymentAmount;
      }

      if (operation.OutstandingBalance is decimal outstandingBalance)
      {
        worksheet.Cell(row, 7).Value = outstandingBalance;
      }

      worksheet.Cell(row, 8).Value = operation.SourceFile;
      if (operation.RemainingInstallmentCount is int remainingInstallments)
      {
        worksheet.Cell(row, 9).Value = remainingInstallments;
      }

      if (operation.RemainingYearsAndMonths is string remainingYearsAndMonths)
      {
        worksheet.Cell(row, 10).Value = remainingYearsAndMonths;
      }
    }

    var lastRow = operations.Count + 1;
    var table = worksheet.Range(1, 1, lastRow, headers.Length).CreateTable("Operations");
    table.Theme = XLTableTheme.TableStyleMedium2;
    worksheet.SheetView.FreezeRows(1);
    worksheet.Column(1).Style.DateFormat.Format = "dd/mm/yyyy";
    worksheet.Column(2).Style.DateFormat.Format = "dd/mm/yyyy";
    worksheet.Column(4).Style.NumberFormat.Format = "#,##0.00";
    worksheet.Column(5).Style.NumberFormat.Format = "#,##0.00";
    worksheet.Column(6).Style.NumberFormat.Format = "#,##0.00";
    worksheet.Column(7).Style.NumberFormat.Format = "#,##0.00";
    worksheet.Columns().AdjustToContents();
    workbook.SaveAs(filePath);
  }
}
