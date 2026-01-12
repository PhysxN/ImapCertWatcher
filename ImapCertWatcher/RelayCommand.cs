using System;
using System.Windows.Input;

namespace ImapCertWatcher.UI
{
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;

        public RelayCommand(Action execute)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        }

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter) => _execute();

        // ✔ Обязательное событие ICommand
        // ✔ Реализовано корректно
        // ✔ Без warning CS0067
        public event EventHandler CanExecuteChanged
        {
            add { }
            remove { }
        }
    }
}
