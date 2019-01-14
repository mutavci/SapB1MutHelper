using System.Collections.Generic;
using SAPbobsCOM;

namespace SapB1MutHelper.Models
{
    public class ChildTable
    {
        public string TableName { get; set; }
        public List<FormColumn> FormColumn { get; set; }
    }

    public class FormColumn
    {
        public string FormColumnAlias { get; set; }
        public string FormColumnDescription { get; set; }
        public BoYesNoEnum Editable { get; set; }
    }
}