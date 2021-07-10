using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Inventor;
using System.Runtime.InteropServices;
using InvAddIn;


namespace InventorAddInNippel
{
    class ButtonMsg
    {

        Inventor.Application m_InvApp;
        //Дефинция кнопки
        Inventor.ButtonDefinition m_ButtonDef;
        //Ribbon-объекты
        ButtonPathToRibbon m_ButtonParhToRibbon;
        //Иконки
        ButtonIcons m_ButtonIcons;
        //ButtonIcons m_ButtonNew;

        Form1 form;

        public ButtonMsg(Inventor.Application InvApp)
        {
            m_InvApp = InvApp;
            //Привязка к событию деактивации AddIn
            StandardAddInServer.RaiseEventAddInDeactivate += OnAddInDeactivate;
            //Создание места на панели куда будет установлена кнопка
            m_ButtonParhToRibbon = new ButtonPathToRibbon(m_InvApp);
            //Создание иконок для кнопки
            m_ButtonIcons = new ButtonIcons(
            InvAddIn.Properties.Resources._01Standart,
            InvAddIn.Properties.Resources.iconBig2, InvAddIn.Properties.Resources.nip);
            //m_ButtonNew = new ButtonIcons(InvAddIn.Properties.Resources.nip,
            //InvAddIn.Properties.Resources.nip);
            //Создание дифинции кнопки и добавление её на панель
            CreateButtonDef_And_InsertToPanel();
        }

        void CreateButtonDef_And_InsertToPanel()
        {
            //Создания дефинции кнопки
            m_ButtonDef = m_InvApp.CommandManager.ControlDefinitions.AddButtonDefinition(
            "Ниппельное соединение", "MyFirstMessageCmd", Inventor.CommandTypesEnum.kQueryOnlyCmdType,
            "", "Ниппельное соединение", "Создать", m_ButtonIcons.IconStandard,
            m_ButtonIcons.IconLarge);

            //Редактирование окна при наведении на кнопку
            m_ButtonDef.ProgressiveToolTip.Title = "khjgjhkgkjhgkjgjgkjhgjhgkhj";
            m_ButtonDef.ProgressiveToolTip.Description = "Проектирование ниппельных соединений при конструировании трупопроводов.";
            m_ButtonDef.ProgressiveToolTip.ExpandedDescription = "Можно проектировать соединения с ниппелями по наружному конусу, шаровыми, коническими. Ниппели и гайки используемые в соединении доступны в Библиотеке компонентов.";
            m_ButtonDef.ProgressiveToolTip.Image = m_ButtonIcons.IconDisplay;

            //Добавление кнопки на панель
            m_ButtonParhToRibbon.RibbonPanelFirst.CommandControls.AddButton(m_ButtonDef,
            true);
            //Привязка к событию клика на кнопку
            m_ButtonDef.OnExecute += OnExecute;
        }

        void OnExecute(Inventor.NameValueMap Context)
        {
            //System.Windows.Forms.MessageBox.Show("Моя первая кнопка");
            form = new Form1();
            form.Show();
        }

        void OnAddInDeactivate()
        {
            //Отключение событий
            m_ButtonDef.OnExecute -= OnExecute;
            StandardAddInServer.RaiseEventAddInDeactivate -= OnAddInDeactivate;
            //Удаление объектов пользовательского интерфейса
            m_ButtonDef.Delete();
            m_ButtonParhToRibbon.RibbonTabFirst.Delete();
        }

        //Контейнер для Ribbon-объектов
        private class ButtonPathToRibbon
        {
            public readonly Ribbon RibbonZeroDoc;
            public readonly RibbonTab RibbonTabFirst;
            public readonly RibbonPanel RibbonPanelFirst;
            //Конструктор
            public ButtonPathToRibbon(Inventor.Application InvApp)
            {
                RibbonZeroDoc = InvApp.UserInterfaceManager.Ribbons["Assembly"];
                RibbonTabFirst = RibbonZeroDoc.RibbonTabs["id_TabDesign"];
                //RibbonPanelFirst = RibbonTabFirst.RibbonPanels.Add("Ниппельные соединения", "id_PanelMessage", "");
                this.RibbonPanelFirst = this.RibbonTabFirst.RibbonPanels["id_PanelA_DesignFasten"];
                //this.RibbonPanelFirst = this.RibbonTabFirst.RibbonPanels["id_PanelA_DesignPowerTransmission"];
            }
        }//Конец класса

        //Контейнер-конвертор для иконок
        private class ButtonIcons
        {
            public readonly stdole.IPictureDisp IconStandard;
            public readonly stdole.IPictureDisp IconLarge;
            public readonly stdole.IPictureDisp IconDisplay;
            //Конструктор
            public ButtonIcons(System.Drawing.Image PictureStandard,
            System.Drawing.Image PictireLarge, System.Drawing.Image PictureDisplay)
            {
                IconStandard = ImageConvertor.ConvertImageToIPictureDisp(PictureStandard);
                IconLarge = ImageConvertor.ConvertImageToIPictureDisp(PictireLarge);
                IconDisplay = ImageConvertor.ConvertImageToIPictureDisp(PictureDisplay);
            }
            private class ImageConvertor : System.Windows.Forms.AxHost
            {

                ImageConvertor() : base("") { }
                public static stdole.IPictureDisp ConvertImageToIPictureDisp(
                System.Drawing.Image Image)
                {
                    if (null == Image) return null;
                    return GetIPictureDispFromPicture(Image) as stdole.IPictureDisp;
                }
            }
        }//конец класса



    }
}
