using System;
using System.Windows;
using System.Windows.Documents;
using Engine.EventArgs;
using Engine.ViewModels;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Session _session;

        public MainWindow()
        {
            InitializeComponent();

            _session = new Session();

            _session.OnMessageRaised += OnMessageRaised;

            DataContext = _session;
        }

        private void btnBillsFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog openFolderDialog = new FolderBrowserDialog();
            openFolderDialog.RootFolder = Environment.SpecialFolder.Desktop;
            openFolderDialog.ShowNewFolderButton = false;
            openFolderDialog.Description = "Please Select The Bill's Folder";
            DialogResult result = openFolderDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                billsPath.Text = String.Empty;
                billsPath.Text = openFolderDialog.SelectedPath;
                _session._billsPath = openFolderDialog.SelectedPath;
            }
            
        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.Title = "Save the resulting file";
            saveFileDialog1.FileName = "Result";
            saveFileDialog1.InitialDirectory = @"C:\";
            DialogResult result = saveFileDialog1.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                if (saveFileDialog1.FileName != "")
                {
                    _session._resultsSaveToPath = saveFileDialog1.FileName;
                    _session.GoButton();
                }
                else
                {
                    _session._resultsSaveToPath = "";
                    _session.GoButton();
                }
            }
        }

        private void btnGo1_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog2 = new SaveFileDialog();
            saveFileDialog2.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog2.DefaultExt = "xlsx";
            saveFileDialog2.Title = "Save the resulting file";
            saveFileDialog2.FileName = "FinalResult";
            saveFileDialog2.InitialDirectory = @"C:\";
            DialogResult result = saveFileDialog2.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                if (saveFileDialog2.FileName != "")
                {
                    _session._resultsSaveToPath1 = saveFileDialog2.FileName;
                    _session.GoButton1();
                }
                else
                {
                    _session._resultsSaveToPath1 = "";
                    _session.GoButton1();
                }
            }
        }

        private void OnMessageRaised(object sender, MessageEventArgs e)
        {
            Messages.Document.Blocks.Add(new Paragraph(new Run(e.Message)));
            Messages.ScrollToEnd();
            Messages1.Document.Blocks.Add(new Paragraph(new Run(e.Message)));
            Messages1.ScrollToEnd();
        }

        private void btnOpenSFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Please Select The S File";
            if (openFileDialog.ShowDialog() == true)
            {
                sPath.Text = String.Empty;
                _session._sPath = openFileDialog.FileName;
                sPath.Text = _session._sPath;
            }
        }

        private void btnOpenPFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Please Select The P File";
            if (openFileDialog.ShowDialog() == true)
                pPath.Text = String.Empty;
            _session._pPath = openFileDialog.FileName;
            pPath.Text = _session._pPath;
        }

        private void btnOpenResultsFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Please Select The Results Weight File";
            if (openFileDialog.ShowDialog() == true)
            {
                resultWeightCostPath.Text = String.Empty;
                _session._resultsPath = openFileDialog.FileName;
                resultWeightCostPath.Text = _session._resultsPath;
            }
        }

        private void btnOpenMaterialAgregadoFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Please Select The 'Material Agregado' File";
            if (openFileDialog.ShowDialog() == true)
                materialAgreadoPath.Text = String.Empty;
            _session._materialAgregadoPath = openFileDialog.FileName;
            materialAgreadoPath.Text = _session._materialAgregadoPath;
        }
    }
}
