using System;
using System.Collections.Generic;
using SapB1MutHelper.Models;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace SapB1MutHelper
{
    public static class MenuCreation
    {

        public static void AddMainSubItems(string ItemId  , List<SideMenu> Menu)
        {
            try
            {
                var oCreationPackage = (MenuCreationParams)Application.SBO_Application.CreateObject(BoCreatableObjectType
                .cot_MenuCreationParams);
                var oMenuItem = Application.SBO_Application.Menus.Item(ItemId);
                var oMenus = oMenuItem.SubMenus;

                foreach (var item in Menu)
                {

                    oCreationPackage.Type = item.Type;
                    oCreationPackage.UniqueID = item.UniqueId;
                    oCreationPackage.String = item.Text;
                    oCreationPackage.Position = -1;
                    oMenus.AddEx(oCreationPackage);

                }

            }
            catch (Exception)
            {
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", BoMessageTime.bmt_Short, true);
            }
        }



    }
}