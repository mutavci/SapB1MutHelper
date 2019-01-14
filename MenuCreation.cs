using System;
using System.Collections.Generic;
using SapB1MutHelper.Models;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace SapB1MutHelper
{
    public static class MenuCreation
    {
        public static void AddMainMenuItems()
        {


            try
            {
                var oCreationPackage = (MenuCreationParams)Application.SBO_Application.CreateObject(BoCreatableObjectType
                    .cot_MenuCreationParams);
                var oMenuItem = Application.SBO_Application.Menus.Item("43520");
                oCreationPackage.Type = BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "HAKT99";
                oCreationPackage.String = "HAKT99";
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = -1;
                oCreationPackage.Image = System.Windows.Forms.Application.StartupPath + @"\pusula.png";
                var oMenus = oMenuItem.SubMenus;
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            {
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", BoMessageTime.bmt_Short);
            }

        }

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
                    oMenus.AddEx(oCreationPackage);


                }


            }
            catch (Exception)
            {
                //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", BoMessageTime.bmt_Short, true);
            }
        }



    }
}