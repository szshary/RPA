using System.IO;
using System.Windows;
//using System.Windows.Shapes;
using Microsoft.Win32;
using RPA.Core;

namespace Rule_File_Selector
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var ruleFileDialog = new OpenFileDialog
            {
                Title = "Lütfen kural dosyasını seçiniz.",
                DefaultExt = ".xml",
                Filter = "XML Dosyaları (*.xml)|*.xml",
                CheckFileExists = true,
                InitialDirectory = @"\\BIM2456\Users\BIM2456\Documents\Visual Studio 2017\Projects\AHE RPA\Automated Processes"
            };

            if (ruleFileDialog.ShowDialog() == true)
            {
                RuleEngine ruleEngine = new RuleEngine(Path.GetDirectoryName(ruleFileDialog.FileName));
                ruleEngine.ProcessTaskFile(ruleFileDialog.SafeFileName);
            }
        }
    }
}