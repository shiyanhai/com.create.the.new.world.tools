using System.Windows;

namespace ExcelToolsGlobalNamingArea
{
    /// <summary>
    /// ExcelTools.xaml Main Logic
    /// </summary>
    public partial class ExcelTools : Window
    {
        private ExcelToolsVM vm = new ExcelToolsVM();

        public ExcelTools()
        {
            InitializeComponent();
            this.DataContext = vm;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtSheetNm.Focus();
        }

        private void BtnBookCreation_Click(object sender, RoutedEventArgs e)
        {
            // Required Items Validation
            if (!this.vm.NotNullValidation())
            {
                MessageBox.Show("Input less than numeric 100 characters, please."
                    , "Warning"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Warning);
                this.txtSheetNo1.Focus();
                return;
            }

            // Single Item Validation
            if (!this.vm.SingleValidation())
            {
                MessageBox.Show("Input less than numeric 100 characters, please."
                    , "Warning"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Warning);
                this.txtSheetNo1.Focus();
                return;
            }

            // Multi Items Validation
            if (!this.vm.MultiValidation())
            {
                MessageBox.Show("The left value should input less than right value, please."
                    , "Warning"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Warning);
                this.txtSheetNo1.Focus();
                return;
            }

            if (!this.vm.IsEnoughSheets())
            {
                MessageBox.Show("Both left and right value should input more than 2, please."
                    , "Warning"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Warning);
                this.txtSheetNo1.Focus();
                return;
            }

            // TODO add Excel interop
            return;

            //new ExcelCreation(vm);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
