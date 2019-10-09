using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MsDyn.Contrib.CloneFieldDefinitions
{
    class ListViewItemComparer: IComparer
    {
        private int col;
        private SortOrder so;
        public ListViewItemComparer()
        {
            col = 0;
        }

        public ListViewItemComparer(int column, SortOrder so)
        {
            col = column;
            this.so = so;
        }

        public int Compare(object x, object y)
        {
            //If sorting by ascending order, compare as usual, 
            //Otherwise invert result
            if(so == SortOrder.Ascending)
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            else return -(String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text));
        }
    }
}
