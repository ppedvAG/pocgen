using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace pocgen.Contracts.Models
{
    public class RelayCommand : ICommand
    {
        private Action execute;
        private Func<bool> canExecute;
        public RelayCommand(Action execute) : this(execute, new Func<bool>(() => true)) {  }
        public RelayCommand(Action execute, Func<bool> canExecute)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }
#pragma warning disable 67
        public event EventHandler CanExecuteChanged;
#pragma warning restore 67
        public bool CanExecute(object parameter) => canExecute?.Invoke() ?? true;
        public void Execute(object parameter = null) => execute?.Invoke();
    }
}
