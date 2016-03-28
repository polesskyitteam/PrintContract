using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Contract
{
    public class MainWindowsView : ViewModelBase
    {
        public string dateContract { get; set; }
        public string name { get; set; }
        public string nameBaby { get; set; }
        public string birthday { get; set; }
        public string addres { get; set; }
        public string phone { get; set; }
        public string eMail { get; set; }
        public string diagnosis { get; set; }
        public string time { get; set; }
        public string serviсe { get; set; }
        public string getResult { get; set; }
        public string many { get; set; }

        private ICommand _print;
        public ICommand Print
        {
            get
            {
                return _print ?? (_print = new RelayCommand(() =>
                    {     
                        PrintContract PrintContract = new PrintContract();
                        PrintContract.WayWord();
                        PrintContract.InsertText(dateContract, name, nameBaby, birthday, addres, phone, eMail, diagnosis, time, serviсe, getResult, many);
                    })) ; 
            }
        }

    }
}
