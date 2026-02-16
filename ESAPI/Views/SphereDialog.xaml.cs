using SFRT_PlanningScript.Models;
using SFRT_PlanningScript.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using VMS.TPS.Common.Model.API;
using ESAPIScript;


namespace SFRT_PlanningScript.Views
{
    /// <summary>
    /// Interaction logic for GridDialog.xaml
    /// </summary>
    /// 
    public partial class SphereDialog : UserControl
    {
        private readonly SphereDialogViewModel vm;
        public TextBoxOutputter outputter;

        public SphereDialog(EsapiWorker EsapiWorker)
        {
            InitializeComponent();
            vm = new SphereDialogViewModel(EsapiWorker);
            DataContext = vm;
        }

        void TimerTick(object state)
        {
            var who = state as string;
            Console.WriteLine(who);
        }

        private void ToggleCircle(object sender, MouseButtonEventArgs e)
        {
            var selectedEllipse = (System.Windows.Shapes.Ellipse)sender;
            Circle selectedCircle = (Circle)selectedEllipse.DataContext;
            selectedCircle.Selected = !selectedCircle.Selected;
        }

        private void CreateLattice(object sender, RoutedEventArgs e)
        {
            vm.CreateLattice();
        }

        private void Optimize(object sender, RoutedEventArgs e)
        {
            vm.Optimize();
        }

        private void AddRoi(object sender, RoutedEventArgs e)
        {
            vm.AddRoi();
        }

        private void RemoveRoi(object sender, RoutedEventArgs e)
        {
            vm.RemoveRoi();
        }




        private void Cancel(object sender, RoutedEventArgs e)
        {
            //this.Close();
        }

    }
}

