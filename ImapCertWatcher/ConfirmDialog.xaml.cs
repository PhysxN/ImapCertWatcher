using System.Windows;
using System.Windows.Input;

namespace ImapCertWatcher.UI
{
    public partial class ConfirmDialog : Window
    {
        public string Message { get; }
        public string TitleText { get; }

        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }

        private ConfirmDialog(string title, string message)
        {
            InitializeComponent();

            TitleText = title;
            Message = message;

            OkCommand = new RelayCommand(() =>
            {
                DialogResult = true;
                Close();
            });

            CancelCommand = new RelayCommand(() =>
            {
                DialogResult = false;
                Close();
            });

            DataContext = this;
        }

        public static bool Show(Window owner, string title, string message)
        {
            var dlg = new ConfirmDialog(title, message)
            {
                Owner = owner
            };
            return dlg.ShowDialog() == true;
        }
    }
}
