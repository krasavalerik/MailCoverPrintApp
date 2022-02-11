using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace THI
{

    [Serializable]
    public class Letter : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _where;
        private string _toWhom;
        private bool _isPrint;

        public string Where {
            get
            {
                return _where;
            }
            set
            {
                _where = value;
                OnPropertyChanged("Where");
            }
        }

        public string ToWhom
        {
            get
            {
                return _toWhom;
            }
            set
            {
                _toWhom = value;
                OnPropertyChanged("ToWhom");
            }
        }
        
        public bool IsPrint
        {
            get
            {
                return _isPrint;
            }
            set
            {
                _isPrint = value;
                OnPropertyChanged("IsPrint");
            }
        }

        public ICommand PreviewCmd
        {
            get { return new PreviewCommand(this); }
        }

        private void Preview()
        {
            var wnd = new EditWindow(this);

            wnd.ShowDialog();
        }

        internal class PreviewCommand : ICommand
        {
            private Letter _letter;

            public PreviewCommand(Letter letter)
            {
                _letter = letter;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object parameter)
            {
                _letter.Preview();
            }
        }
    }
}
