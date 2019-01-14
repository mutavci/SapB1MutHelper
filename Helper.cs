using System;
using SAPbouiCOM;
using Company = SAPbobsCOM.Company;

namespace SapB1MutHelper
{
    public static class Helper
    {
        public static Company OCompany { get; set; }

        public static bool IsDtExist(this IForm oForm, string check)
        {
            DataTable checkDt;
            try
            {
                checkDt = oForm.DataSources.DataTables.Item(check);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public static bool IsCflExist(this IForm oForm, string check)
        {
            ChooseFromList checkCfl;
            try
            {
                checkCfl = oForm.ChooseFromLists.Item(check);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public static bool IsItemExist(this IForm oForm, string check)
        {
            Item checkItem;
            try
            {
                checkItem = oForm.Items.Item(check);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
    }
}