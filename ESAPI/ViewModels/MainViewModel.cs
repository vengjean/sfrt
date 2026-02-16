using Prism.Mvvm;

namespace SFRT_PlanningScript.ViewModels
{
    public class MainViewModel : BindableBase
    {
        private string postText;
        public string PostText
        {
            get { return postText; }
            set { SetProperty(ref postText, value); }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri)
             );
            e.Handled = true;
        }

        public MainViewModel()
        {
            //var isDebug = SFRT_PlanningScript.Properties.Settings.Default.Debug;
            ////MessageBox.Show($"Display Terms {isDebug}");
            //PostText = "";
            //if (isDebug) { PostText += " *** Not Validated For Clinical Use ***"; }
        }
    }
}
