using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeLib {
    public class PPTInfo {

        public string fullFilename;
        public string singleFilename;
        public string abstruct;
        public string createTime;
        public string title;
        public string presenter;

        public string[] titles;
        public string[] note;
        public string[] jpgName;
        public string[] jpgURL;

        public int count;


        public PPTInfo(int num) {

            this.count = num;
            titles = new string[num];

            note = new string[num];
            jpgName = new string[num];
            jpgURL = new string[num];


        }

    }
}
