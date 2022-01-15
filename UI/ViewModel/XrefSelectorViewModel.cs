using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace EpicLumExporter.UI.ViewModel
{
    public class XrefSelectorViewModel : INPC
    {
        public List<string> xrefs { get; set; }
        public int selectedIndex { get; set; }

        public ICommand btnOK { get; set; }

        public event EventHandler OnRequestClose;

        public XrefSelectorViewModel()
        {
            xrefs = new List<string>();
            selectedIndex = -1;
            btnOK = new RCommand(btnOKexecute);
        }

        private void btnOKexecute(object obj)
        {
            OnRequestClose(this, new EventArgs());
        }
    }
}
