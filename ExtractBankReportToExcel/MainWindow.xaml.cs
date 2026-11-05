using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using UglyToad.PdfPig.Core;

namespace ExtractBankReportToExcel;

public partial class MainWindow : Window
{
  private readonly string _pdfFolder = Path.Combine(AppContext.BaseDirectory, "PdfFiles");

  public ObservableCollection<BankOperation> Operations { get; } = [];

  public MainWindow()
  {
    InitializeComponent();
    DataContext = this;
    CreatePdfFileDirectory();
    AddVersionTotitle();
    FolderPathText.Text = $"Dossier des PDF : {_pdfFolder}";
    Loaded += async (_, _) => await LoadPdfReportsAsync();
  }

  private void AddVersionTotitle()
  {
    var version = Assembly.GetExecutingAssembly().GetName().Version;
    if (version != null)
    {
      Title += $" V{version}";
    }
  }

  private static void CreatePdfFileDirectory()
  {
    try
    {
      Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "PdfFiles"));
    }
    catch (IOException exception)
    {
      ShowDirectoryError(exception.Message);
      throw;
    }
    catch (UnauthorizedAccessException exception)
    {
      ShowDirectoryError(exception.Message);
      throw;
    }
  }

  private static void ShowDirectoryError(string message)
  {
    MessageBox.Show($"Impossible de créer le dossier PdfFiles : {message}", "Erreur de création du dossier", MessageBoxButton.OK, MessageBoxImage.Error);
  }

  private async void RefreshButton_Click(object sender, RoutedEventArgs e)
  {
    await LoadPdfReportsAsync();
  }

  private async Task LoadPdfReportsAsync()
  {
    SetBusy(true);
    Operations.Clear();

    try
    {
      var pdfFiles = Directory.GetFiles(_pdfFolder, "*.pdf", SearchOption.TopDirectoryOnly)
        .OrderBy(path => path, StringComparer.CurrentCultureIgnoreCase)
        .ToArray();

      if (pdfFiles.Length == 0)
      {
        StatusText.Text = "Aucun fichier PDF trouvé. Copiez vos relevés dans le dossier indiqué ci-dessus.";
        return;
      }

      var result = await Task.Run(() => ReadPdfFiles(pdfFiles));
      foreach (var operation in result.Operations)
      {
        Operations.Add(operation);
      }

      ExportButton.IsEnabled = Operations.Count > 0;
      StatusText.Text = $"{Operations.Count} opération(s) extraite(s) depuis {pdfFiles.Length} fichier(s) PDF.";

      if (result.Warnings.Count > 0)
      {
        var warningMessage = string.Join(Environment.NewLine, result.Warnings);
        StatusText.Text += $" {result.Warnings.Count} avertissement(s).";
        MessageBox.Show(warningMessage, "Certains relevés nécessitent une vérification", MessageBoxButton.OK, MessageBoxImage.Warning);
      }
    }
    catch (IOException exception)
    {
      ShowLoadError(exception.Message);
    }
    catch (UnauthorizedAccessException exception)
    {
      ShowLoadError(exception.Message);
    }
    finally
    {
      SetBusy(false);
    }
  }

  private static (List<BankOperation> Operations, List<string> Warnings) ReadPdfFiles(IEnumerable<string> pdfFiles)
  {
    var operations = new List<BankOperation>();
    var warnings = new List<string>();

    foreach (var pdfFile in pdfFiles)
    {
      try
      {
        var fileOperations = PdfReportReader.Read(pdfFile);
        if (fileOperations.Count == 0)
        {
          warnings.Add($"{Path.GetFileName(pdfFile)} : aucune opération reconnue dans ce relevé.");
          continue;
        }

        if (fileOperations[0].MonthlyPaymentAmount is null)
        {
          warnings.Add($"{Path.GetFileName(pdfFile)} : montant du prochain prélèvement mensuel introuvable.");
        }

        if (fileOperations[0].OutstandingBalance is null)
        {
          warnings.Add($"{Path.GetFileName(pdfFile)} : montant restant dû des opérations en plusieurs fois introuvable.");
        }

        if (fileOperations[0].InstallmentMonthlyPaymentAmount is null || fileOperations[0].InstallmentMonthlyPaymentAmount <= 0)
        {
          warnings.Add($"{Path.GetFileName(pdfFile)} : mensualité des opérations en plusieurs fois introuvable ou invalide ; durée restante non calculée.");
        }

        operations.AddRange(fileOperations);
      }
      catch (IOException exception)
      {
        warnings.Add($"{Path.GetFileName(pdfFile)} : {exception.Message}");
      }
      catch (UnauthorizedAccessException exception)
      {
        warnings.Add($"{Path.GetFileName(pdfFile)} : {exception.Message}");
      }
      catch (ArgumentException exception)
      {
        warnings.Add($"{Path.GetFileName(pdfFile)} : {exception.Message}");
      }
      catch (PdfDocumentFormatException exception)
      {
        warnings.Add($"{Path.GetFileName(pdfFile)} : PDF illisible ({exception.Message})");
      }
    }

    return (operations, warnings);
  }

  private void ExportButton_Click(object sender, RoutedEventArgs e)
  {
    if (Operations.Count == 0)
    {
      return;
    }

    var dialog = new SaveFileDialog
    {
      FileName = $"Releves_bancaires_{DateTime.Now:yyyyMMdd}.xlsx",
      Filter = "Classeur Excel (*.xlsx)|*.xlsx",
      DefaultExt = ".xlsx",
      AddExtension = true
    };

    if (dialog.ShowDialog(this) != true)
    {
      return;
    }

    try
    {
      ExcelReportExporter.Export(dialog.FileName, Operations.ToArray());
      StatusText.Text = $"Export terminé : {dialog.FileName}";
      MessageBox.Show(
        "Le fichier Excel a été créé.",
        "Export terminé",
        MessageBoxButton.OK,
        MessageBoxImage.Information);
    }
    catch (IOException exception)
    {
      ShowExportError(exception.Message);
    }
    catch (UnauthorizedAccessException exception)
    {
      ShowExportError(exception.Message);
    }
  }

  private void SetBusy(bool isBusy)
  {
    RefreshButton.IsEnabled = !isBusy;
    ExportButton.IsEnabled = !isBusy && Operations.Count > 0;
    Mouse.OverrideCursor = isBusy ? Cursors.Wait : null;
  }

  private void ShowLoadError(string message)
  {
    StatusText.Text = "Le chargement des PDF a échoué.";
    MessageBox.Show(message, "Erreur de lecture", MessageBoxButton.OK, MessageBoxImage.Error);
  }

  private static void ShowExportError(string message)
  {
    MessageBox.Show(message, "Erreur d’export", MessageBoxButton.OK, MessageBoxImage.Error);
  }
}
