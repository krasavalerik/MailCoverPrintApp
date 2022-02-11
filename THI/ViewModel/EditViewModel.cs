using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace THI
{
    public class EditViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public EditViewModel(Letter ltr)
        {
            _letter = ltr;
        }
        
        private Letter _letter;

        public Letter Letter
        {
            get
            {
                return _letter;
            }
            set
            {
                _letter = value;
                OnPropertyChanged("Letter");

            }
        }
    }
}
