using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;
using System.Text;
using Microsoft.Win32;

namespace StructX
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string _excelPath;
        public MainWindow() => InitializeComponent();

        private void BtnGenerarSql_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "SQL Script (*.sql)|*.sql",
                Title = "Guardar script SQL",
                FileName = System.IO.Path.GetFileNameWithoutExtension(_excelPath) + ".sql"
            };

            if (saveDialog.ShowDialog() == true)
            {
                string sql = GenerarSqlDesdeExcel(_excelPath);
                File.WriteAllText(saveDialog.FileName, sql);

                MessageBox.Show("Script SQL generado",
                    "StructX",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }


        private void BtnCargarExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "Selecciona un archivo Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                _excelPath = dialog.FileName;

                TxtFileName.Text = System.IO.Path.GetFileName(_excelPath);

                FilePreview.Visibility = Visibility.Visible;
                BtnGenerarSql.Visibility = Visibility.Visible;
                TxtEmpty.Visibility = Visibility.Collapsed;
            }
        }

    private string GenerarSqlDesdeExcel(string excelPath)
    {
        var sb = new StringBuilder();

        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet(1);

        string tableName = System.IO.Path.GetFileNameWithoutExtension(excelPath);

        sb.AppendLine($"CREATE TABLE [{tableName}] (");

        var headers = ws.Row(1).CellsUsed();

        foreach (var cell in headers)
        {
            string columnName = cell.GetString().Replace(" ", "_");
            sb.AppendLine($"    [{columnName}] VARCHAR(255),");
        }

        sb.Length -= 3;
        sb.AppendLine("\n);");

        return sb.ToString();
    }


}
}