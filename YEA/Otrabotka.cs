using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YEA
{
    public class Otrabotka : Oblozhka
    {
        private List<ListViewItem> dataCollection;
       

        public Otrabotka(List<ListViewItem> data, string surname) : base(surname)
        {
            dataCollection = data;
        }

        public Otrabotka()
        {

        }

        public List<ListViewItem> DataCollection
        {
            get { return dataCollection; }
            set { dataCollection = value; }
        }

       
    }

}
