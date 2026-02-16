using SFRT_PlanningScript.Models;
using SFRT_PlanningScript.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using VMS.TPS.Common.Model.API;

namespace SFRT_PlanningScript.Views
{
    /// <summary>
    /// Interaction logic for GridMLCDialog.xaml
    /// </summary>
    public partial class GridMLCDialog : UserControl
    {
        public GridMLCDialogViewModel vm;

        public GridMLCDialog(ScriptContext context)
        {
            InitializeComponent();
            vm = new GridMLCDialogViewModel(context);
            this.DataContext = vm;
        }

        private void CreateGrid(object sender, RoutedEventArgs e)
        {
            vm.CreateGrid();
        }

        private void Cancel(object sender, RoutedEventArgs e)
        {
        }
    }
}
