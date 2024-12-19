using System.Windows;

namespace CSVProcessorApp
{
    public partial class GratuityInputWindow : Window
    {
        public double GratuityPercentage { get; private set; } = 0.0;

        public GratuityInputWindow()
        {
            InitializeComponent();
        }

        private void OnSubmitClick(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(GratuityPercentageTextBox.Text, out double gratuity) && gratuity >= 0)
            {
                GratuityPercentage = gratuity;
                this.DialogResult = true; // Indicates success
                this.Close();
            }
            else
            {
                MessageBox.Show("Please enter a valid gratuity percentage (0 or higher).", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
