using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace pocgen.Contracts.Models
{
    public class BaseViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public virtual Dispatcher DispatcherObject { get; protected set; }

        protected BaseViewModel()
        {
            DispatcherObject = Dispatcher.CurrentDispatcher;
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        protected virtual bool SetValue<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            MessagingCenter.Send(this,"Log", new LoggerEventArgs(propertyName, MethodBase.GetCurrentMethod().Name,$"Value changed to '{ value?.ToString() ?? "null"}'"));
            return true;
        }
    }
}
