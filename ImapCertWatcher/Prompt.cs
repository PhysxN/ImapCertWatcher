using System.Windows;
using System.Windows.Controls;

namespace ImapCertWatcher
{
    public static class Prompt
    {
        public static string ShowDialog(string text, string caption)
        {
            var win = new Window
            {
                Title = caption,
                Width = 420,
                Height = 170,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };

            if (Application.Current != null && Application.Current.MainWindow != null)
                win.Owner = Application.Current.MainWindow;

            var root = new Grid
            {
                Margin = new Thickness(12)
            };

            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var txtLabel = new TextBlock
            {
                Text = text ?? "",
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(txtLabel, 0);
            root.Children.Add(txtLabel);

            var tb = new TextBox
            {
                MinWidth = 360,
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(tb, 1);
            root.Children.Add(tb);

            var buttonsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var btnOk = new Button
            {
                Content = "OK",
                Width = 80,
                IsDefault = true,
                Margin = new Thickness(0, 0, 8, 0)
            };
            btnOk.Click += (_, __) => win.DialogResult = true;

            var btnCancel = new Button
            {
                Content = "Отмена",
                Width = 80,
                IsCancel = true
            };
            btnCancel.Click += (_, __) => win.DialogResult = false;

            buttonsPanel.Children.Add(btnOk);
            buttonsPanel.Children.Add(btnCancel);

            Grid.SetRow(buttonsPanel, 2);
            root.Children.Add(buttonsPanel);

            win.Content = root;
            win.Loaded += (_, __) => tb.Focus();

            return win.ShowDialog() == true ? tb.Text : null;
        }
    }
}