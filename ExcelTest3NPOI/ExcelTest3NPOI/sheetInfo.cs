using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTest3NPOI
{
    public class sheetInfo
    {
        private int rowMax;
        private int listMax;
        private int[] rowIndex = new int[10];
        #region getter,setter
        public int RowMax
        {
            get
            {
                return rowMax;
            }

            set
            {
                rowMax = value;
            }
        }

        public int ListMax
        {
            get
            {
                return listMax;
            }

            set
            {
                listMax = value;
            }
        }

        public int[] RowIndex
        {
            get
            {
                return rowIndex;
            }

            set
            {
                rowIndex = value;
            }
        }


        #endregion
    }
}
