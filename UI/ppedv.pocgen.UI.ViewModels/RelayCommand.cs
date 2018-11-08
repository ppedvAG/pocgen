﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Input;

namespace ppedv.pocgen.UI.ViewModels
{
    public class RelayCommand : ICommand
    {
        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            this.execute = execute;
            this.canExecute = canExecute ?? new Func<object, bool>(arg => true);
        }

        private readonly Action<object> execute;
        private readonly Func<object, bool> canExecute;

#pragma warning disable 67
        public event EventHandler CanExecuteChanged;
#pragma warning restore 67
        public bool CanExecute(object parameter = null) => canExecute.Invoke(parameter);
        public void Execute(object parameter = null) => execute?.Invoke(parameter);
    }
}