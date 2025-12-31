using Autodesk.Revit.UI;
using System.Windows;
using ValveFlangeMulti.UI.ViewModels;

namespace ValveFlangeMulti.UI
{
    public partial class ValveFlangeMultiWindow : Window
    {
        public ValveFlangeMultiWindow(ExternalCommandData commandData)
        {
            InitializeComponent();
            DataContext = new MainViewModel(commandData);
        }
    }
}
