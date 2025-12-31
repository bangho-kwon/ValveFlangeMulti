using Autodesk.Revit.UI;
using System;
using System.Windows;
using ValveFlangeMulti.UI.ViewModels;

namespace ValveFlangeMulti.UI
{
    public partial class ValveFlangeMultiWindow : Window
    {
        public ValveFlangeMultiWindow(ExternalCommandData commandData)
        {
            try
            {
                // Validate input
                if (commandData == null)
                    throw new ArgumentNullException(nameof(commandData), "ExternalCommandData cannot be null");

                // Initialize UI components
                InitializeComponent();

                // Create and set the view model (constructor will throw if fails)
                DataContext = new MainViewModel(commandData);
            }
            catch (Exception ex)
            {
                // Show error to user
                MessageBox.Show(
                    $"Failed to initialize ValveFlangeMulti window:\n\n{ex.Message}\n\nType: {ex.GetType().Name}",
                    "Initialization Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                
                // Re-throw to let caller handle
                throw;
            }
        }
    }
}
