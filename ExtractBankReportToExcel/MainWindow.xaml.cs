using System.IO;
using System.Windows;

namespace ExtractBankReportToExcel
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow: Window
  {
    public MainWindow()
    {
      InitializeComponent();
      CreatePdfFileDirectory();
    }

    private static void CreatePdfFileDirectory()
    {
      if (!Directory.Exists("PdfFiles"))
      {
        try
        {
          Directory.CreateDirectory("PdfFiles");
        }
        catch (Exception exception)
        {
          MessageBox.Show($"Could not create PdfFiles directory: {exception.Message}", "Can't create directory", MessageBoxButton.OK, MessageBoxImage.Error);
          throw;
        }
      }
    }
  }
}