using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;


namespace THI
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private ObservableCollection<Letter> _table = new ObservableCollection<Letter>();
        private bool _isPrintAll;

        public ObservableCollection<Letter> Table
        {
            get
            {
                return _table;
            }
            set
            {
                _table = value;
                OnPropertyChanged("Table");

            }
        }

        public bool isPrintAll
        {
            get
            {
                return _isPrintAll;
            }
            set
            {
                _isPrintAll = value;
                OnPropertyChanged("isPrintAll");

                _table.ToList().ForEach(x => x.IsPrint = value);
            }
        } 

        public ICommand FromExcellCmd
        {
            get { return new FromExcellCommand(this); }
        }

        public ICommand SaveCmd
        {
            get { return new SaveCommand(this); }
        }

        public ICommand LoadCmd
        {
            get { return new LoadCommand(this); }
        }
        public ICommand ExitCmd
        {
            get { return new ExitCommand(this); }
        }

        public ICommand PrintCmd
        {
            get { return new PrintCommand(this); }
        }

        public ICommand AddCmd
        {
            get { return new AddCommand(this); }
        }

        public ICommand DeleteCmd
        {
            get { return new DeleteCommand(this); }
        }

        private void LoadFromExcell()
        {
            OpenFileDialog openDialog = new OpenFileDialog();

            if (openDialog.ShowDialog().Value)
            {
                Excel.Application obj = new Excel.Application();
                Excel.Workbook book = obj.Workbooks.Open(openDialog.FileName);
                Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[1];

                var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);


                for (int i = 2; i <= lastCell.Row; i++)
                {
                    var letter = new Letter();

                    letter.ToWhom = sheet.Cells[i, 1].Text;
                    letter.Where = sheet.Cells[i, 2].Text;

                    _table.Add(letter);
                }

                book.Close();                
            }
        }

        private void Load()
        {
            OpenFileDialog openDialog = new OpenFileDialog();

            if (openDialog.ShowDialog().Value)
            {
                string fname = openDialog.FileName;

                openDialog.DefaultExt = "xml";

                Stream stream = File.OpenRead(fname);
                XmlSerializer serials = new XmlSerializer(typeof(ObservableCollection<Letter>));                                               

                Table = (ObservableCollection<Letter>)serials.Deserialize(stream);

                stream.Close();
            }
        }

        private void Save()
        {
            SaveFileDialog dlg = new SaveFileDialog();
            
            
            if (dlg.ShowDialog() == true)
            {
                string fname = dlg.FileName + ".xml";              

                Stream stream = File.Create(fname);

                XmlSerializer serial = new XmlSerializer(typeof(ObservableCollection<Letter>));
                serial.Serialize(stream, _table);
                stream.Close();
            }
        }

        private void Exit(object wnd)
        {
            ((Window)wnd).Close();
        }


        private void Print()
        {

            FixedDocument fixedDoc = new FixedDocument();
            PageContent pageContent = new PageContent();
            FixedPage fixedPage = new FixedPage();

            StackPanel sp = new StackPanel();

            var printTbl = _table.ToList().FindAll(x => x.IsPrint == true);


            foreach (var vr in printTbl)
            {
                var ctrl = new Cover();
                ctrl.DataContext = new EditViewModel(vr);


                sp.Children.Add(ctrl);

            }

            fixedPage.Children.Add(sp);

            ((System.Windows.Markup.IAddChild)pageContent).AddChild(fixedPage);
            fixedDoc.Pages.Add(pageContent);



            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                printDialog.PrintDocument(((IDocumentPaginatorSource)fixedDoc).DocumentPaginator, "Распечатываем");
            }

        }

        private void AddLetter()
        {
            Letter let = new Letter();

            EditWindow wnd = new EditWindow(let);
            wnd.ShowDialog();

            _table.Add(let);
        }

        private void DeleteLetter()
        {
            var deleteTbl = _table.ToList().FindAll(x=>x.IsPrint==true);

            foreach (var vr in deleteTbl)
                _table.Remove(vr);
        }

        internal class FromExcellCommand : ICommand
        {
            private MainViewModel _vm;

            public FromExcellCommand(MainViewModel letter)
            {
                _vm = letter;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object parameter)
            {
                _vm.LoadFromExcell();
            }
        }

        internal class SaveCommand : ICommand
        {
            private MainViewModel _vm;

            public SaveCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object parameter)
            {
                _vm.Save();
            }
        }

        internal class LoadCommand : ICommand
        {
            private MainViewModel _vm;

            public LoadCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object parameter)
            {
                _vm.Load();
            }
        }
        
        internal class ExitCommand : ICommand
        {
            private MainViewModel _vm;

            public ExitCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object wnd)
            {
                _vm.Exit(wnd);
            }
        }

        internal class PrintCommand : ICommand
        {
            private MainViewModel _vm;

            public PrintCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object wnd)
            {
                _vm.Print();
            }
        }

        internal class AddCommand : ICommand
        {
            private MainViewModel _vm;

            public AddCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parametr)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object wnd)
            {
                _vm.AddLetter();
            }
        }

        internal class DeleteCommand : ICommand
        {
            private MainViewModel _vm;

            public DeleteCommand(MainViewModel vm)
            {
                _vm = vm;
            }

            public bool CanExecute(object parametr)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged
                = delegate { };

            public void Execute(object wnd)
            {
                _vm.DeleteLetter();
            }
        }
    }
}
