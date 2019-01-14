using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;

namespace SapB1MutHelper.Models
{

    //Sap B1 Yan Menu ve Alt Menü
   public class SideMenu
    {
        public BoMenuType Type { get; set; }
        public string UniqueId { get; set; }
        public string Text { get; set; }
        public string Image { get; set; }
        public string Pozition { get; set; }
    }

}
