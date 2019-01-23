# SapB1MutHelper
SAP business one yardımcısı Tablo,Udo,Triger,Menü Oluşturucu


## Nasıl Kullanılır ?
Adım :1
Öncelikli olarak Visual studio üzerinden SAP Add-On Projesi başlatınız

Adım :2 
NuGet Gallery üzerinden SapB1MutHelper yazarak yükleyebilirsiniz ya da PM kullanarak 

```
Install-Package SapB1MutHelper -Version 1.0.0
```
devam edebilirsiniz 

Adım :3
program.cs dosyasının Main(){} fonksiyonunun içerisine 

```
Helper.OCompany = (Company)Application.SBO_Application.Company.GetDICompany(); 
```
kodunu ekleyiniz burada Arkada çalışan Sap b1'ın hangi firmada(bildiğimiz database e verdikleri isim) calışıyorsa o nesneyi çekiyoruz
ve ardından Ana Bir Menü Ekleyelim

```
                var MainMenuList = new List<SideMenu>();
                var MainMenu = new SideMenu
                {
                    UniqueId = "MAINMENUID",
                    Type = BoMenuType.mt_POPUP,
                    Text = "Main Menüm",
                    Image = System.Windows.Forms.Application.StartupPath + @"\pusula.png",
                    Pozition = "-1"
                };

                MainMenuList.Add(MainMenu);
```
