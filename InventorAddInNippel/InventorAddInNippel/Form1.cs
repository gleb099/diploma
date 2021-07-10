using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//Библиотеки для Inventor
using Inventor;
using System.Diagnostics;
using System.Reflection;

namespace InvAddIn
{
    public partial class Form1 : Form
    {
        //////////////////////////////////////////////////////// Мусор
        private void button1_Click(object sender, EventArgs e)
        {
            //Кнопка "Открыть Inventor"

            System.Windows.Forms.MessageBox.Show("Запуск Инвентора может занять несколько минут!");
            System.Diagnostics.Process.Start("C:/Program Files/Autodesk/Inventor 2021/Bin/Inventor.exe");
            //label2.Text = "Инвентор открыт!";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Кнопка "Собрать"
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Кнопка "Test"
        }

        private void button30_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Кнопка "Подобрать"
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Pass mf
        }

        //Выбор типа соединения this combobox was deleted
        //private void comboBox15_TextChanged(object sender, EventArgs e)
        //{
        //    if (comboBox15.Text == "С шаровым ниппелем под сварку")
        //    {
        //        tabControl1.SelectedTab = tabPage1;
        //    }
        //    else if (comboBox15.Text == "С шаровым ниппелем под пайку")
        //    {
        //        tabControl1.SelectedTab = tabPage2;
        //    }
        //    else if (comboBox15.Text == "С коническим ниппелем исп. 1")
        //    {
        //        tabControl1.SelectedTab = tabPage3;
        //    }
        //    else if (comboBox15.Text == "С коническим ниппелем исп. 2")
        //    {
        //        tabControl1.SelectedTab = tabPage4;
        //    }
        //    else if (comboBox15.Text == "По конусу 1")
        //    {
        //        tabControl1.SelectedTab = tabPage5;
        //    }
        //    else if (comboBox15.Text == "По конусу 2")
        //    {
        //        tabControl1.SelectedTab = tabPage6;
        //    }
        //}
        ////////////////////////////////////////////////////////

        /// <summary>
        /// ThisApplication - Объект для определения активного состояния Инвентора
        /// </summary>
        private Inventor.Application ThisApplication = null;
        /// <summary>
        /// Словарь для хранения ссылок на документы деталей
        /// </summary>
        private Dictionary<string, PartDocument> oPartDoc = new Dictionary<string, PartDocument>();
        /// <summary>
        /// Словарь для хранения ссылок на определения деталей
        /// </summary>
        private Dictionary<string, PartComponentDefinition> oCompDef = new Dictionary<string, PartComponentDefinition>();
        /// <summary>
        /// Словарь для хранения ссылок на инструменты создания деталей
        /// </summary>
        private Dictionary<string, TransientGeometry> oTransGeom = new Dictionary<string, TransientGeometry>();
        /// <summary>
        /// Словарь для хранения ссылок на транзакции редактирования
        /// </summary>
        private Dictionary<string, Transaction> oTrans = new Dictionary<string, Transaction>();
        /// <summary>
        /// Словарь для хранения имен сохраненных документов деталей
        /// </summary>
        private Dictionary<string, string> oFileName = new Dictionary<string, string>();
        private AssemblyDocument oAssemblyDocName;

        //TabControl1.ItemSize = new Size(0, 1);

        static UserInputEvents UIevents;
        static UserInputEvents UIevents1;
        static UserInputEvents UIevents2;
        [DllImport("User32.Dll")]
        private static extern bool GetAsyncKeyState(int vKey);
        const int LB = 1;
        const int RB = 2;
        int clickedItemN1 = 0;
        //Для кликов на обычные ребра/грани
        public static int flagG = 0;
        public static int flagP = 0;
        public static int flagK1 = 0;
        public static int flagK2 = 0;
        public static int flagKDn1 = 0;
        public static int flagKDn = 0;
        public static int flagPNK1 = 0;
        public static int flagPNK2 = 0;
        //Для клика на грань с резьбой
        public static int flagGR = 0;
        public static int flagPR = 0;
        public static int flagK1R = 0;
        public static int flagK2R = 0;
        public static int flagPNK1R = 0;
        public static int flagPNK2R = 0;

        public static Face faceK;

        //Соединение с шаровым ниппелем исполнение 1
        public static string categNameGM = "ГАЙКИ 23353";
        public static string categNameNip23 = "Шаровой ниппель исполнение 1";

        public static string g23 = "ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ";
        public static string g23RowName = "ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-04";

        public static string nip23 = "ГОСТ 23355-78 Исполнение 1";
        public static string nip23RowName = "ГОСТ 23355-78 Исполнение 1-05";

        //Соединение с шаровым ниппелем исполнение 2
        public static string categNameNipSH2 = "Шаровой ниппель исполнение 2";
        public static string nipSH2 = "ГОСТ 23355-78 Исполнение 2";
        public static string nipSH2RowName = "ГОСТ 23355-78 Исполнение 2-05";

        //Соединение с коническим ниппелем исполнение 1
        public static string categNameKonNip1 = "Конический ниппель исполнение 2";
        public static string konNip1 = "ГОСТ 28016-89 Исполнение 1";
        public static string konNip1RowName = "ГОСТ 28016-89 Исполнение 1-05";

        //Соединение с коническим ниппелем исполнение 2
        public static string categNameKonNip2 = "Конический ниппель исполнение 2";
        public static string konNip2 = "ГОСТ 28016-89 Исполнение_2";
        public static string konNip2RowName = "ГОСТ 28016-89 Исполнение_2-05";

        public static string categNameRing = "Конический ниппель исполнение 1";
        public static string ring = "Кольцо для соединения 9833";
        public static string ringRowName = "005-007-14";
        //Соединение с ниппелем по наружному конусу исполнение 1
        public static string categNameGM1 = "ГАЙКИ 13957-74";
        public static string categNameNipPK1 = "Ниппель по конусу исполнение 1";

        public static string g2 = "Гайки 13957-74";
        public static string g2RowName = "ГОСТ 13957-74 ПО НАРУЖНОМУ КОНУСУ-01";

        public static string nipPK1 = "ГОСТ 13956-74 Исполнение 1";
        public static string nipPK1RowName = "ГОСТ 13956-74 Исполнение 1-01";

        //Соединение с ниппелем по наружному конусу исполнение 2
        public static string categNameNipPK2 = "Ниппель по конусу исполнение 2";
        public static string nipPK2 = "ГОСТ 13956-74 Исполнение 2";
        public static string nipPK2RowName = "ГОСТ 13956-74 Исполнение 1-01";

        //Поиск фейсов для вытсавления зависимостей в сборках
        //По наружному конусу 2
        public static Face facePNK2Kas;
        public static Face facePNK2Os;
        public static bool flagPNK2faceKas = false;
        public static bool flagPNK2faceOs = false;

        //По наружному конусу 1
        public static Face facePNK1Kas;
        public static Face facePNK1Os;
        public static bool flagPNK1faceKas = false;
        public static bool flagPNK1faceOs = false;

        //Конический ниппель с углом 24 исполнение 2
        public static Face faceKon2Vs;
        public static Face faceKon2Os;
        public static bool flagKon2faceVs = false;
        public static bool flagKon2faceOs = false;

        //Конический ниппель с углом 24 исполнение 1
        public static Face faceKon1Vs;
        public static Face faceKon1Os;
        public static bool flagKon1faceVs = false;
        public static bool flagKon1faceOs = false;

        //Шаровой ниппель на пайке (исполнение 2)
        public static Face faceSh2Vs;
        public static bool flagSh2faceVs = false;

        //Шаровой ниппель под сварку (исполнение 1)
        public static Face faceSh1Vs;
        public static bool flagSh1faceVs = false;

        //---------------------------Трубы---------------------------
        public static Face faceTr1;
        public static bool faceTr1bl = false;

        public static Face faceTr2;
        public static bool faceTr2bl = false;

        public static Face faceTr3;
        public static bool faceTr3bl = false;

        public static Face faceTr4;
        public static bool faceTr4bl = false;

        public static Face faceTr5;
        public static bool faceTr5bl = false;

        public static Face faceTr6;
        public static bool faceTr6bl = false;

        public Form1()
        {
            InitializeComponent();
            try
            {
                //Проверка наличия активного состояния Инвентора.
                ThisApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Запуск Инвентора может занять несколько минут!");
                System.Diagnostics.Process.Start("C:/Program Files/Autodesk/Inventor 2021/Bin/Inventor.exe");
            }

            //Получение объекта-Инвентора invApp - для чека клика и поиска выбранной грани
            //Inventor.Application invApp = ThisApplication;
            //UIevents = invApp.CommandManager.UserInputEvents;
            ////Подписка на событие
            //UIevents.OnSelect += UIevents_OnSelect;
            ////System.Windows.Forms.MessageBox.Show("SasNo");
            ////System.Console.Read();

            //Для кликов
            Inventor.Application invApp = ThisApplication;
            File file = invApp.ActiveDocument.File;
            UIevents = invApp.CommandManager.UserInputEvents;

            //Работа с ComboBox и TabControl
            tabControl1.ItemSize = new Size(0, 1);
            //comboBox15.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //comboBox15.Text = "С шаровым ниппелем под сварку";



            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000;
            toolTip1.ShowAlways = true;
            toolTip1.ToolTipTitle = "Ребро Dn";
            toolTip1.SetToolTip(this.checkBox2, "Выбрать ребро для добавления диаметра Dn");
            toolTip1.SetToolTip(this.checkBox16, "Выбрать ребро для добавления диаметра Dn");
            toolTip1.SetToolTip(this.checkBox20, "Выбрать ребро для добавления диаметра Dn");
            toolTip1.SetToolTip(this.checkBox24, "Выбрать ребро для добавления диаметра Dn");
            toolTip1.SetToolTip(this.checkBox28, "Выбрать ребро для добавления диаметра Dn");
            toolTip1.SetToolTip(this.checkBox31, "Выбрать ребро для добавления диаметра Dn");

            ToolTip toolTip2 = new ToolTip();
            toolTip2.AutoPopDelay = 5000;
            toolTip2.ShowAlways = true;
            toolTip2.ToolTipTitle = "Ребро Dv";
            toolTip2.SetToolTip(this.checkBox34, "Выбрать ребро для добавления диаметра Dv");
            toolTip2.SetToolTip(this.checkBox33, "Выбрать ребро для добавления диаметра Dv");
            toolTip2.SetToolTip(this.checkBox19, "Выбрать ребро для добавления диаметра Dv");
            toolTip2.SetToolTip(this.checkBox23, "Выбрать ребро для добавления диаметра Dv");

            ToolTip toolTip3 = new ToolTip();
            toolTip3.AutoPopDelay = 5000;
            toolTip3.ShowAlways = true;
            toolTip3.ToolTipTitle = "Резьба";
            toolTip3.SetToolTip(this.checkBox1, "Выбрать грань с резьбой для добавления обозначения");
            toolTip3.SetToolTip(this.checkBox17, "Выбрать грань с резьбой для добавления обозначения");
            toolTip3.SetToolTip(this.checkBox21, "Выбрать грань с резьбой для добавления обозначения");
            toolTip3.SetToolTip(this.checkBox25, "Выбрать грань с резьбой для добавления обозначения");
        }

        //Функция, которая найдет ветку категории в библиотеке компонентов
        public static ContentTreeViewNode GetCategoryNode(Inventor.Application inventorapp, string categName1)
        {
            var enumerator = inventorapp.ContentCenter.TreeViewTopNode.ChildNodes;
            ContentTreeViewNode gostLibNode = null;
            foreach (ContentTreeViewNode node in enumerator)
            {
                if (node.DisplayName == categName1)
                {
                    gostLibNode = node;
                    break;
                }
            }
            return gostLibNode;
        }

        //Функция, которая найдет конктреное семейство в библиотеке компонентов
        public static ContentFamily GetContentFamily(ContentFamiliesEnumerator families, string FamilyName)
        {
            ContentFamily contentFamily = null;
            foreach (ContentFamily family in families)
            {
                if (family.DisplayName == FamilyName)
                {
                    contentFamily = family;
                    break;
                }
            }
            return contentFamily;
        }

        //Функция, которая найдет строку исполнения семейства параметрической детали в библиотеке компонентов
        public static ContentTableRow GetTablelRow(ContentTableRows rows, string searchColumn, string searchValue)
        {
            ContentTableRow contentTableRow = null;
            foreach (ContentTableRow row in rows)
            {
                if (row[searchColumn].Value == searchValue)
                {
                    contentTableRow = row;
                    break;
                }

            }
            if (contentTableRow == null)
            {
                MessageBox.Show($"Строка {searchValue} не найдена");
            }
            return contentTableRow;
        }

        //Функция, которая выбирает исполнение детали по названию
        public void changeIsp(iPartFactory partFactory, string nameIsp)
        {
            iPartTableRows rows = partFactory.TableRows;
            Console.WriteLine("iPart");
            foreach (iPartTableRow row in rows)
            {
                if (row[1].Value == nameIsp)
                {
                    partFactory.DefaultRow = row;
                }
                Console.WriteLine(row[1].Value);
            }
            Console.WriteLine("iPart end");
        }

        //Функция, которая включит режим "Выбор граней и ребер" в сборке
        public void chooseSelection(bool choose)
        {
            AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
            SelectionPriorityEnum d = SelectionPriorityEnum.kEdgeAndFaceSelectionPriority;
            SelectionPriorityEnum d1 = SelectionPriorityEnum.kComponentSelectionPriority;
            if (choose == true)
            {
                oAssDoc.SelectionPriority = d;
            }
            else if (choose == false)
            {
                oAssDoc.SelectionPriority = d1;
            }
        }

        //Тестовая функция для определения граней и ребер в сборке на клик по кнопке
        private void button23_Click(object sender, EventArgs e)
        {
            chooseSelection(true);
        }

        //Определяем по клику Наружный диаметр

        //Dn 2

        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            //условный проход
            if (checkBox24.Checked == true)
            {
                checkBox23.Checked = false;
                checkBox25.Checked = false;
                checkBox26.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagKDn = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventskDn_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox24.Checked == false)
            {
                flagKDn = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventskDn_OnSelect;
                //UIevents.OnSelect -= UIeventskDn_OnSelect;
                UIevents.OnSelect -= UIeventskDn_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventskDn_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox14.Text = Convert.ToString(dy);
                    UIevents.OnSelect -= UIeventskDn_OnSelect;
                    checkBox24.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventskDn_OnSelect;
                    checkBox24.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventskDn_OnSelect;
                checkBox24.Checked = false;
            }
        }



        //Dn 1
        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox22.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagKDn1 = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventskDn1_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox20.Checked == false)
            {
                flagKDn1 = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventskDn1_OnSelect;
                //UIevents.OnSelect -= UIeventskDn1_OnSelect;
                UIevents.OnSelect -= UIeventskDn1_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventskDn1_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox3.Text = Convert.ToString(dy);
                    UIevents.OnSelect -= UIeventskDn1_OnSelect;
                    checkBox20.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventskDn1_OnSelect;
                    checkBox20.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventskDn1_OnSelect;
                checkBox20.Checked = false;
            }
        }


        // Определяем по клику Условный проход

        //C шаровым ниппелем под сварку
        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox15.Checked = false;
                checkBox34.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagG = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsG_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox2.Checked == false)
            {
                flagG = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsG_OnSelect;
                //UIevents.OnSelect -= UIeventsG_OnSelect;
                UIevents.OnSelect -= UIeventsG_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsG_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;

                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox10.Text = Convert.ToString(dy);
                    UIevents.OnSelect -= UIeventsG_OnSelect;
                    checkBox2.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsG_OnSelect;
                    checkBox2.Checked = false;
                }

            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsG_OnSelect;
                checkBox2.Checked = false;
            }

        }


        private void checkBox34_CheckedChanged(object sender, EventArgs e)
        {
            //UIevents.OnSelect -= UIeventsG_OnSelect;
            Inventor.Application invApp = ThisApplication;
            File file = invApp.ActiveDocument.File;
            UIevents2 = invApp.CommandManager.UserInputEvents;
            
            if (checkBox34.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox1.Checked = false;
                checkBox15.Checked = false;
                UIevents2.OnSelect += UIevents2_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox34.Checked == false)
            {
                //UIevents2.OnSelect += UIevents2_OnSelect;
                //UIevents2.OnSelect -= UIevents2_OnSelect;
                UIevents2.OnSelect -= UIevents2_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIevents2_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dn = d.Radius * 10 * 2;
                    comboBox16.Text = Convert.ToString(dn);
                    UIevents2.OnSelect -= UIevents2_OnSelect;
                    checkBox34.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents2.OnSelect -= UIevents2_OnSelect;
                    checkBox34.Checked = false;
                }

            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents2.OnSelect -= UIevents2_OnSelect;
                checkBox34.Checked = false;
            }
        }

        //C шаровым ниппелем на пайке
        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox33.Checked = false;
                flagP = 1;
                //stopSub(true);
                UIevents.OnSelect += UIeventsP_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox16.Checked == false)
            {
                flagP = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsP_OnSelect;
                //UIevents.OnSelect -= UIeventsP_OnSelect;
                //MessageBox.Show("SasNo");
                UIevents.OnSelect -= UIeventsP_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsP_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox11.Text = Convert.ToString(dy);
                    //View.Fit(); //приближает деталь
                    UIevents.OnSelect -= UIeventsP_OnSelect;
                    checkBox16.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsP_OnSelect;
                    checkBox16.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsP_OnSelect;
                checkBox16.Checked = false;
            }
        }

        private void checkBox33_CheckedChanged(object sender, EventArgs e)
        {
            Inventor.Application invApp = ThisApplication;
            File file = invApp.ActiveDocument.File;
            UIevents1 = invApp.CommandManager.UserInputEvents;
            //Подписка на событие
            if (checkBox33.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox16.Checked = false;
                UIevents1.OnSelect += UIevents1_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox33.Checked == false)
            {
                //UIevents1.OnSelect += UIevents1_OnSelect;
                //UIevents1.OnSelect -= UIevents1_OnSelect;
                //MessageBox.Show("SasNo");
                UIevents1.OnSelect -= UIevents1_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIevents1_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dn = d.Radius * 10 * 2;
                    comboBox15.Text = Convert.ToString(dn);
                    checkBox33.Checked = false;
                    UIevents1.OnSelect -= UIevents1_OnSelect;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    checkBox33.Checked = false;
                    UIevents1.OnSelect -= UIevents1_OnSelect;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                checkBox33.Checked = false;
                UIevents1.OnSelect -= UIevents1_OnSelect;
            }
        }


            //С коническим ниппелем исполнение 1
        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox19.Checked == true)
            {
                checkBox20.Checked = false;
                checkBox21.Checked = false;
                checkBox22.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagK1 = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsK1_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox19.Checked == false)
            {
                flagK1 = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsK1_OnSelect;
                //UIevents.OnSelect -= UIeventsK1_OnSelect;
                UIevents.OnSelect -= UIeventsK1_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsK1_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox12.Text = Convert.ToString(dy);
                    UIevents.OnSelect -= UIeventsK1_OnSelect;
                    checkBox19.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsK1_OnSelect;
                    checkBox19.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsK1_OnSelect;
                checkBox19.Checked = false;
            }
        }


        //С коническим ниппелем исполнение 2
        private void button14_Click(object sender, EventArgs e)
        {

        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox23.Checked == true)
            {
                checkBox24.Checked = false;
                checkBox25.Checked = false;
                checkBox26.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagK2 = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsK2_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox23.Checked == false)
            {
                flagK2 = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsK2_OnSelect;
                //UIevents.OnSelect -= UIeventsK2_OnSelect;
                UIevents.OnSelect -= UIeventsK2_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsK2_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dy = d.Radius * 10 * 2;
                    comboBox13.Text = Convert.ToString(dy);
                    UIevents.OnSelect -= UIeventsK2_OnSelect;
                    checkBox23.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsK2_OnSelect;
                    checkBox23.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsK2_OnSelect;
                checkBox23.Checked = false;
            }
        }

        //По наружному конусу исполнение 1
        private void button15_Click(object sender, EventArgs e)
        {

        }

        private void checkBox28_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox28.Checked == true)
            {
                checkBox27.Checked = false;
                checkBox29.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagPNK1 = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsPNK1_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox28.Checked == false)
            {
                flagPNK1 = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPNK1_OnSelect;
                //UIevents.OnSelect -= UIeventsPNK1_OnSelect;
                UIevents.OnSelect -= UIeventsPNK1_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPNK1_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dn = d.Radius * 10 * 2;
                    comboBox6.Text = Convert.ToString(dn);
                    UIevents.OnSelect -= UIeventsPNK1_OnSelect;
                    checkBox28.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsPNK1_OnSelect;
                    checkBox28.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsPNK1_OnSelect;
                checkBox28.Checked = false;
            }
        }

        //По наружному конусу исполнение 2
        private void button16_Click(object sender, EventArgs e)
        {

        }

        private void checkBox31_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox31.Checked == true)
            {
                checkBox30.Checked = false;
                checkBox32.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagPNK2 = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsPNK2_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox31.Checked == false)
            {
                flagPNK2 = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPNK2_OnSelect;
                //UIevents.OnSelect -= UIeventsPNK2_OnSelect;
                UIevents.OnSelect -= UIeventsPNK2_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPNK2_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Edge)
            {
                Inventor.Edge obj1 = obj as Inventor.Edge;
                if (Convert.ToString(obj1.CurveType) == "kCircleCurve")
                {
                    CurveTypeEnum temp = obj1.CurveType;
                    Circle d = obj1.Curve[temp] as Circle;
                    double dn = d.Radius * 10 * 2;
                    comboBox8.Text = Convert.ToString(dn);
                    UIevents.OnSelect -= UIeventsPNK2_OnSelect;
                    checkBox31.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsPNK2_OnSelect;
                    checkBox31.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете окружность", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsPNK2_OnSelect;
                checkBox31.Checked = false;
            }
        }

        // Определяем по клику Резьбу

        //С шаровым ниппелем под сварку
        private void button17_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox15.Checked = false;
                checkBox34.Checked = false;

                UIevents.OnSelect += UIeventsGR_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox1.Checked == false)
            {

                UIevents.OnSelect -= UIeventsGR_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsGR_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {

                Inventor.Face obj1 = obj as Inventor.Face;
                if (Convert.ToString(obj1.ThreadInfos) != "")
                {
                    ThreadInfo r = obj1.ThreadInfos as ThreadInfo;
                    foreach (ThreadInfo temp in obj1.ThreadInfos)
                    {
                        r = temp;
                    }
                    comboBox1.Text = Convert.ToString(r.ThreadDesignation);
                    UIevents.OnSelect -= UIeventsGR_OnSelect;
                    checkBox1.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsGR_OnSelect;
                    checkBox1.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsGR_OnSelect;
                checkBox1.Checked = false;
            }
        }

        //С шаровым ниппелем под пайку
        private void button18_Click(object sender, EventArgs e)
        {

        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                checkBox16.Checked = false;
                checkBox18.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagPR = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsPR_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox17.Checked == false)
            {
                flagPR = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPR_OnSelect;
                //UIevents.OnSelect -= UIeventsPR_OnSelect;
                UIevents.OnSelect -= UIeventsPR_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPR_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;
                if (Convert.ToString(obj1.ThreadInfos) != "")
                {
                    ThreadInfo r = obj1.ThreadInfos as ThreadInfo;
                    foreach (ThreadInfo temp in obj1.ThreadInfos)
                    {
                        r = temp;
                    }
                    comboBox4.Text = Convert.ToString(r.ThreadDesignation);
                    UIevents.OnSelect -= UIeventsPR_OnSelect;
                    checkBox17.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsPR_OnSelect;
                    checkBox17.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsPR_OnSelect;
                checkBox17.Checked = false;
            }
        }

        //С коническим ниппелем исполнение 1
        private void button19_Click(object sender, EventArgs e)
        {

        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox21.Checked == true)
            {
                checkBox20.Checked = false;
                checkBox19.Checked = false;
                checkBox22.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagK1R = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsK1R_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox21.Checked == false)
            {
                flagK1R = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsK1R_OnSelect;
                //UIevents.OnSelect -= UIeventsK1R_OnSelect;
                UIevents.OnSelect -= UIeventsK1R_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsK1R_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;
                if (Convert.ToString(obj1.ThreadInfos) != "")
                {
                    ThreadInfo r = obj1.ThreadInfos as ThreadInfo;
                    foreach (ThreadInfo temp in obj1.ThreadInfos)
                    {
                        r = temp;
                    }
                    comboBox2.Text = Convert.ToString(r.ThreadDesignation);
                    UIevents.OnSelect -= UIeventsK1R_OnSelect;
                    checkBox21.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsK1R_OnSelect;
                    checkBox21.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsK1R_OnSelect;
                checkBox21.Checked = false;
            }
        }

        //С коническим ниппелем исполнение 2
        private void button20_Click(object sender, EventArgs e)
        {

        }

        private void checkBox25_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox25.Checked == true)
            {
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox26.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagK2R = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsK2R_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox25.Checked == false)
            {
                flagK2R = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsK2R_OnSelect;
                //UIevents.OnSelect -= UIeventsK2R_OnSelect;
                UIevents.OnSelect -= UIeventsK2R_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsK2R_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;
                if (Convert.ToString(obj1.ThreadInfos) != "")
                {
                    ThreadInfo r = obj1.ThreadInfos as ThreadInfo;
                    foreach (ThreadInfo temp in obj1.ThreadInfos)
                    {
                        r = temp;
                    }
                    comboBox5.Text = Convert.ToString(r.ThreadDesignation);
                    UIevents.OnSelect -= UIeventsK2R_OnSelect;
                    checkBox25.Checked = false;
                }
                else
                {
                    MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsK2R_OnSelect;
                    checkBox25.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете грань с резьбой", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsK2R_OnSelect;
                checkBox25.Checked = false;
            }
        }

        //По наружному конусу исполнение 1
        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void checkBox27_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox27.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox15.Checked = false;
                Inventor.Application invApp = ThisApplication;
                File file = invApp.ActiveDocument.File;
                flagPNK1R = 1;

                //stopSub(true);
                UIevents.OnSelect += UIeventsPNK1R_OnSelect;
                chooseSelection(true);
            }
            else if (checkBox27.Checked == false)
            {
                flagPNK1R = 0;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPNK1R_OnSelect;
                //UIevents.OnSelect -= UIeventsPNK1R_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPNK1R_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;
                ThreadInfo r = obj1.ThreadInfos as ThreadInfo;
                foreach (ThreadInfo temp in obj1.ThreadInfos)
                {
                    r = temp;
                }
                comboBox7.Text = Convert.ToString(r.ThreadDesignation);
                UIevents.OnSelect -= UIeventsPNK1R_OnSelect;
                checkBox27.Checked = false;
                //View.Fit(); //приближает деталь
            }
        }

        //По наружному конусу исполнение 2
        private void button22_Click(object sender, EventArgs e)
        {

        }

        //Поиск фейсов для вставления зависимостей в сборке

        //Поиск фейсов в "по наружному конусу 2"
        private void button24_Click(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                flagPNK2faceKas = true;
            }
            else
            {
                faceTr6bl = true;
            }
        }

        private void checkBox32_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox32.Checked == true)
            {
                checkBox30.Checked = false;
                checkBox31.Checked = false;
                if (checkBox14.Checked == true)
                {
                    //stopSub(true);
                    flagPNK2faceKas = true;
                    UIevents.OnSelect += UIeventsPNK2face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox13.Checked == true)
                {
                    //stopSub(true);
                    faceTr6bl = true;
                    UIevents.OnSelect += UIeventsPNK2face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox14.Checked == checkBox13.Checked == false)
                {
                    //stopSub(false);
                    flagPNK2faceKas = false;
                    faceTr6bl = false;
                    chooseSelection(false);
                    checkBox32.Checked = false;
                    //UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                    UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox32.Checked == false)
            {
                flagPNK2faceKas = false;
                faceTr6bl = false;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPNK2face_OnSelect;
                //UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPNK2face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;

                if ((Convert.ToString(obj1.SurfaceType) == "kConeSurface") && (obj1.Edges.Count == 2))
                {
                    facePNK2Kas = obj as Face;
                    faceTr6 = obj as Face;
                    button12.Enabled = true;
                    UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                    checkBox32.Checked = false;
                    MessageBox.Show($"Грань выбрана");
                }
                else
                {
                    MessageBox.Show("Выберете внешнюю коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                    checkBox32.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете внешнюю коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsPNK2face_OnSelect;
                checkBox32.Checked = false;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            flagPNK2faceOs = true;
        }

        //Поиск фейсов в "по наружному конусу 1"
        private void button26_Click(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                flagPNK1faceKas = true;
            }
            else
            {
                faceTr5bl = true;
            }
        }

        private void checkBox29_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox29.Checked == true)
            {
                checkBox27.Checked = false;
                checkBox28.Checked = false;
                if (checkBox12.Checked == true)
                {
                    flagPNK1faceKas = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsPNK1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox11.Checked == true)
                {
                    faceTr5bl = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsPNK1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox12.Checked == checkBox11.Checked == false)
                {
                    flagPNK1faceKas = false;
                    faceTr5bl = false;
                    chooseSelection(false);
                    checkBox29.Checked = false;
                    //stopSub(false);
                    UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox29.Checked == false)
            {
                flagPNK1faceKas = false;
                faceTr5bl = false;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsPNK1face_OnSelect;
                //UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsPNK1face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {

                Inventor.Face obj1 = obj as Inventor.Face;

                if ((Convert.ToString(obj1.SurfaceType) == "kConeSurface") && (obj1.Edges.Count == 2))
                {
                    facePNK1Kas = obj as Face;
                    faceTr5 = obj as Face;
                    button11.Enabled = true;
                    UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                    checkBox29.Checked = false;
                    MessageBox.Show($"Грань выбрана");
                }
                else
                {
                    MessageBox.Show("Выберете внешнюю коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                    checkBox29.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете внешнюю коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsPNK1face_OnSelect;
                checkBox29.Checked = false;
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            flagPNK1faceOs = true;
        }

        //Поиск фейсов в "с коническим ниппелем 2"
        private void button28_Click(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                flagKon2faceVs = true;
            }
            else
            {
                faceTr4bl = true;
            }
        }

        private void checkBox26_CheckedChanged(object sender, EventArgs e)
        {
            //comboBox15.Text = comboBox5.Text;

            if (checkBox26.Checked == true)
            {
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox25.Checked = false;
                if (checkBox10.Checked == true)
                {
                    flagKon2faceVs = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsKon2face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox9.Checked == true)
                {
                    faceTr4bl = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsKon2face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox10.Checked == checkBox9.Checked == false)
                {
                    flagKon2faceVs = false;
                    faceTr4bl = false;
                    chooseSelection(false);
                    checkBox26.Checked = false;
                    //stopSub(false);
                    //UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                    UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox26.Checked == false)
            {
                flagKon2faceVs = false;
                faceTr4bl = false;
                //stopSub(true);
                //UIevents.OnSelect += UIeventsKon2face_OnSelect;
                //UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsKon2face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;

                if (checkBox9.Checked == true)
                {
                    if ((Convert.ToString(obj1.SurfaceType) == "kPlaneSurface") && (obj1.Edges.Count == 2))
                    {
                        faceKon2Vs = obj as Face;
                        faceTr4 = obj as Face;
                        button10.Enabled = true;
                        UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                        checkBox26.Checked = false;
                        MessageBox.Show($"Грань выбрана");
                    }
                    else
                    {
                        MessageBox.Show("Выберете плоскую грань в виде кольца", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                        checkBox26.Checked = false;
                    }
                }

                else if (checkBox10.Checked == true)
                {
                    if ((Convert.ToString(obj1.SurfaceType) == "kConeSurface") && (obj1.Edges.Count == 2))
                    {
                        faceKon2Vs = obj as Face;
                        faceTr4 = obj as Face;
                        button10.Enabled = true;
                        UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                        checkBox26.Checked = false;
                        MessageBox.Show($"Грань выбрана");
                    }
                    else
                    {
                        MessageBox.Show("Выберете коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                        checkBox26.Checked = false;
                    }
                }
            }

            else
            {
                if (checkBox10.Checked == true)
                {
                    MessageBox.Show("Выберете коническую грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                    checkBox26.Checked = false;
                }
                else if (checkBox9.Checked == true)
                {
                    MessageBox.Show("Выберете плоскую грань в виде кольца", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsKon2face_OnSelect;
                    checkBox26.Checked = false;
                }

            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            flagKon2faceOs = true;
        }

        //Поиск фейсов в "с коническим ниппелем 1"
        private void button32_Click(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                flagKon1faceVs = true;
            }
            else
            {
                faceTr3bl = true;
            }
        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox22.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox20.Checked = false;
                checkBox21.Checked = false;
                if (checkBox8.Checked == true)
                {
                    flagKon1faceVs = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsKon1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox7.Checked == true)
                {
                    faceTr3bl = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsKon1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox8.Checked == checkBox7.Checked == false)
                {
                    flagKon1faceVs = false;
                    faceTr3bl = false;
                    chooseSelection(false);
                    checkBox22.Checked = false;
                    //stopSub(false);
                    //UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox22.Checked == false)
            {
                flagKon1faceVs = false;
                faceTr3bl = false;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsKon1face_OnSelect;
                //UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                chooseSelection(false);
                
            }
        }

        private void UIeventsKon1face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;

                if (checkBox7.Checked == true)
                {
                    if ((Convert.ToString(obj1.SurfaceType) == "kPlaneSurface") && (obj1.Edges.Count == 2))
                    {
                        faceKon1Vs = obj as Face;
                        faceTr3 = obj as Face;
                        button7.Enabled = true;
                        UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                        checkBox22.Checked = false;
                        MessageBox.Show($"Грань выбрана");
                    }
                    else
                    {
                        MessageBox.Show("Выберете плоскую грань в виде кольца", 
                            "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                        checkBox22.Checked = false;
                    }
                }

                else if (checkBox8.Checked == true)
                {
                    if ((Convert.ToString(obj1.SurfaceType) == "kConeSurface") && (obj1.Edges.Count == 2))
                    {
                        faceKon1Vs = obj as Face;
                        faceTr3 = obj as Face;
                        button7.Enabled = true;
                        UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                        checkBox22.Checked = false;
                        MessageBox.Show($"Грань выбрана");
                    }
                    else
                    {
                        MessageBox.Show("Выберете коническую грань", 
                            "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                        checkBox22.Checked = false;
                    }
                }
            }

            else
            {
                if (checkBox8.Checked == true)
                {
                    MessageBox.Show("Выберете грань", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                    checkBox22.Checked = false;
                }
                else if (checkBox7.Checked == true)
                {
                    MessageBox.Show("Выберете плоскую грань в виде кольца", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsKon1face_OnSelect;
                    checkBox22.Checked = false;
                }
                   
            }
        }


        private void button33_Click(object sender, EventArgs e)
        {
            flagKon1faceOs = true;
        }

        //Поиск фейсов в "с шаровым ниппелем на пайке"
        private void button34_Click(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                flagSh2faceVs = true;
            }
            else
            {
                faceTr2bl = true;
            }
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                checkBox16.Checked = false;
                checkBox17.Checked = false;
                checkBox33.Checked = false;
                if (checkBox6.Checked == true)
                {
                    flagSh2faceVs = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsSh2face_OnSelect;
                    chooseSelection(true);
                    
                }
                else if (checkBox5.Checked == true)
                {
                    faceTr2bl = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsSh2face_OnSelect;
                    chooseSelection(true);
                   
                }
                else if (checkBox6.Checked == checkBox5.Checked == false)
                {
                    flagSh2faceVs = false;
                    faceTr2bl = false;
                    chooseSelection(false);
                    checkBox18.Checked = false;
                    //stopSub(false);
                    UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox18.Checked == false)
            {
                flagSh2faceVs = false;
                faceTr2bl = false;
                //stopSub(false);
                //UIevents.OnSelect += UIeventsSh2face_OnSelect;
                //UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                chooseSelection(false);
                
            }
        }

        private void UIeventsSh2face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;

                if ((Convert.ToString(obj1.SurfaceType) == "kPlaneSurface") && (obj1.Edges.Count == 2))
                {
                    faceSh2Vs = obj as Face;
                    faceTr2 = obj as Face;
                    button8.Enabled = true;
                    UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                    checkBox18.Checked = false;
                    MessageBox.Show($"Грань выбрана");
                }
                else
                {
                    MessageBox.Show("Выберете плоскую грань в виде кольца", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                    checkBox18.Checked = false;
                }
            }
            else
            {
                MessageBox.Show("Выберете плоскую грань в виде кольца", "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsSh2face_OnSelect;
                checkBox18.Checked = false;
            }
            
        }

        //Поиск фейсов в "с шаровым ниппелем под сварку"
        private void button35_Click(object sender, EventArgs e)
        {
            flagSh1faceVs = true;
        }

        //------------------------------------Для труб-------------------------------------

        //Грань для вставки в трубу (шаровой ниппель под сварку)
        private void button36_Click(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                flagSh1faceVs = true;
            }
            else
            {
                faceTr1bl = true;
            }
        }
        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox15.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox1.Checked = false;
                if (checkBox4.Checked == true)
                {
                    flagSh1faceVs = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsSh1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox3.Checked == true)
                {
                    faceTr1bl = true;
                    //stopSub(true);
                    UIevents.OnSelect += UIeventsSh1face_OnSelect;
                    chooseSelection(true);
                }
                else if (checkBox4.Checked == checkBox3.Checked == false)
                {
                    flagSh1faceVs = false;
                    faceTr1bl = false;
                    chooseSelection(false);
                    checkBox15.Checked = false;
                    //stopSub(false);
                    //UIevents.OnSelect -= UIeventsSh1face_OnSelect;
                    MessageBox.Show($"Не выбран тип");
                }
            }
            else if (checkBox15.Checked == false)
            {
                flagSh1faceVs = false;
                faceTr1bl = false;
                UIevents.OnSelect -= UIeventsSh1face_OnSelect;
                chooseSelection(false);
            }
        }

        private void UIeventsSh1face_OnSelect(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;


                if ((Convert.ToString(obj1.SurfaceType) == "kPlaneSurface") && (obj1.Edges.Count == 2))
                {
                    faceSh1Vs = obj as Face;
                    faceTr1 = obj as Face;
                    button5.Enabled = true;
                    UIevents.OnSelect -= UIeventsSh1face_OnSelect;
                    checkBox15.Checked = false;
                    MessageBox.Show($"Грань выбрана");
                } 
                else
                {
                    MessageBox.Show("Выберете плоскую грань в виде кольца", 
                        "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UIevents.OnSelect -= UIeventsSh1face_OnSelect;
                    checkBox15.Checked = false;
                }
                
            }
            else
            {
                MessageBox.Show("Выберете плоскую грань в виде кольца", 
                    "Ошибка при выборе параметра", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UIevents.OnSelect -= UIeventsSh1face_OnSelect;
                checkBox15.Checked = false;
            }
        }


        //Чекбоксы на тип соединения
        private void checkBox4_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox4.Checked == true || checkBox3.Checked == true)
            {
                comboBox10.Enabled = true;
                checkBox2.Enabled = true;
            }

            if (checkBox4.Checked == true)
            {
                checkBox3.Checked = false;
                pictureBox1.Image = global::InvAddIn.Properties.Resources.sh1_sh;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox15, "Выбрать плоскую грань на штуцере для добавления соединения в сборку");
            }
            else
            {
                checkBox3.Checked = true;
                pictureBox1.Image = global::InvAddIn.Properties.Resources.sh1_tr;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox15, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
        }

        private void checkBox3_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox4.Checked == true || checkBox3.Checked == true)
            {
                comboBox10.Enabled = true;
                checkBox2.Enabled = true;
            }

            if (checkBox3.Checked == true)
            {
                checkBox4.Checked = false;
                pictureBox1.Image = global::InvAddIn.Properties.Resources.sh1_tr;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox15, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
            else
            {
                checkBox4.Checked = true;
                pictureBox1.Image = global::InvAddIn.Properties.Resources.sh1_sh;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox15, "Выбрать плоскую грань на штуцере для добавления соединения в сборку");
            }
        }

        //Грань для вставки в трубу (шаровой ниппель под пайку)

        //Чекбоксы на тип соединения
        private void checkBox6_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox6.Checked == true || checkBox5.Checked == true)
            {
                comboBox11.Enabled = true;
                checkBox16.Enabled = true;
            }

            if (checkBox6.Checked == true)
            {
                checkBox5.Checked = false;
                pictureBox2.Image = global::InvAddIn.Properties.Resources.sh2_sh;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox18, "Выбрать плоскую грань на штуцере для добавления соединения в сборку");
            }
            else
            {
                checkBox5.Checked = true;
                pictureBox2.Image = global::InvAddIn.Properties.Resources.sh2_tr;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox18, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
        }

        private void checkBox5_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox6.Checked == true || checkBox5.Checked == true)
            {
                comboBox11.Enabled = true;
                checkBox16.Enabled = true;
            }

            if (checkBox5.Checked == true)
            {
                checkBox6.Checked = false;
                pictureBox2.Image = global::InvAddIn.Properties.Resources.sh2_tr;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox18, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
            else
            {
                checkBox6.Checked = true;
                pictureBox2.Image = global::InvAddIn.Properties.Resources.sh2_sh;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox18, "Выбрать плоскую грань на штуцере для добавления соединения в сборку");
            }
        }

        //Грань для вставки в трубу(конический ниппель 1)

        //Чекбоксы на тип соединения
        private void checkBox8_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox7.Checked == true || checkBox8.Checked == true)
            {
                comboBox3.Enabled = true;
                checkBox20.Enabled = true;
            }

            if (checkBox8.Checked == true)
            {
                checkBox7.Checked = false;
                pictureBox3.Image = global::InvAddIn.Properties.Resources.Н_КН1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox22, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
            else
            {
                checkBox7.Checked = true;
                pictureBox3.Image = global::InvAddIn.Properties.Resources.Н_КН12;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox22, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
        }

        private void checkBox7_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox7.Checked == true || checkBox8.Checked == true)
            {
                comboBox3.Enabled = true;
                checkBox20.Enabled = true;
            }

            if (checkBox7.Checked == true)
            {
                checkBox8.Checked = false;
                pictureBox3.Image = global::InvAddIn.Properties.Resources.Н_КН12;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox22, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
            else
            {
                checkBox8.Checked = true;
                pictureBox3.Image = global::InvAddIn.Properties.Resources.Н_КН1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox22, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
        }

        //Грань для вставки в трубу(кониечский ниппель 2)
        private void checkBox10_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox10.Checked == true || checkBox9.Checked == true)
            {
                comboBox14.Enabled = true;
                checkBox24.Enabled = true;
            }

            if (checkBox10.Checked == true)
            {
                checkBox9.Checked = false;
                pictureBox4.Image = global::InvAddIn.Properties.Resources.Н_КН22;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox26, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
            else
            {
                checkBox9.Checked = true;
                pictureBox4.Image = global::InvAddIn.Properties.Resources.Н_КН2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox26, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
        }

        private void checkBox9_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox10.Checked == true || checkBox9.Checked == true)
            {
                comboBox14.Enabled = true;
                checkBox24.Enabled = true;
            }

            if (checkBox9.Checked == true)
            {
                checkBox10.Checked = false;
                pictureBox4.Image = global::InvAddIn.Properties.Resources.Н_КН2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Плоская грань";
                toolTip1.SetToolTip(this.checkBox26, "Выбрать плоскую грань на трубе для добавления соединения в сборку");
            }
            else
            {
                checkBox10.Checked = true;
                pictureBox4.Image = global::InvAddIn.Properties.Resources.Н_КН22;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox26, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
        }

        //Грань для вставки в трубу(по конусу 1)
        private void checkBox12_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox12.Checked == true || checkBox11.Checked == true)
            {
                comboBox6.Enabled = true;
                checkBox28.Enabled = true;
            }

            if (checkBox12.Checked == true)
            {
                checkBox11.Checked = false;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox29, "Выбрать коническую грань штуцера для добавления соединения в сборку");

            }
            else
            {
                checkBox11.Checked = true;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox29, "Выбрать внешнюю коническую грань развальцованной трубы для добавления соединения в сборку");
            }
        }

        private void checkBox11_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox12.Checked == true || checkBox11.Checked == true)
            {
                comboBox6.Enabled = true;
                checkBox28.Enabled = true;
            }

            if (checkBox11.Checked == true)
            {
                checkBox12.Checked = false;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox29, "Выбрать внешнюю коническую грань развальцованной трубы для добавления соединения в сборку");
            }
            else
            {
                checkBox12.Checked = true;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК1;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox29, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
        }

        //Грань для вставки в трубу(по конусу 2)

        private void checkBox14_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox14.Checked == true || checkBox13.Checked == true)
            {
                comboBox8.Enabled = true;
                checkBox31.Enabled = true;
            }

            if (checkBox14.Checked == true)
            {
                checkBox13.Checked = false;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox32, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
            else
            {
                checkBox13.Checked = true;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox32, "Выбрать внешнюю коническую грань развальцованной трубы для добавления соединения в сборку");
            }
        }

        private void checkBox13_Click(object sender, EventArgs e)
        {
            //Включаю элементы управления
            if (checkBox14.Checked == true || checkBox13.Checked == true)
            {
                comboBox8.Enabled = true;
                checkBox31.Enabled = true;
            }

            if (checkBox13.Checked == true)
            {
                checkBox14.Checked = false;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox32, "Выбрать внешнюю коническую грань развальцованной трубы для добавления соединения в сборку");
            }
            else
            {
                checkBox14.Checked = true;
                pictureBox5.Image = global::InvAddIn.Properties.Resources.Н_ПК2;
                ToolTip toolTip1 = new ToolTip();
                toolTip1.AutoPopDelay = 5000;
                toolTip1.ShowAlways = true;
                toolTip1.ToolTipTitle = "Коническая грань";
                toolTip1.SetToolTip(this.checkBox32, "Выбрать коническую грань штуцера для добавления соединения в сборку");
            }
        }


        //Инфа и сборочки

        //Соединение с шаровым ниппелем под сварку

        //Изменение резьбы
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            string rez = "";
            string r1 = "";
            rez = comboBox1.Text;
            r1 = comboBox10.Text;
            //comboBox1.Update();
            //comboBox10.Update();
            bool temp1 = false;
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                if (Convert.ToString(rez) == Convert.ToString(comboBox1.Items[i]))
                {
                    checkBox15.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                checkBox15.Enabled = false;
            }

        }

        //Изменение условного прохода
        private void comboBox10_TextChanged(object sender, EventArgs e)
        {
            string rez = "";
            string r1 = "";
            rez = comboBox1.Text;
            r1 = comboBox10.Text;
            //comboBox1.Update();
            //comboBox10.Update();
            bool temp1 = false;
            for (int i = 0; i < comboBox10.Items.Count; i++)
            {
                if (Convert.ToString(r1) == Convert.ToString(comboBox10.Items[i]))
                {
                    if (checkBox4.Checked == true)
                    {
                        comboBox1.Enabled = true;
                        checkBox1.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    else if (checkBox3.Checked == true)
                    {
                        comboBox16.Enabled = true;
                        checkBox34.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    
                }
            }
            if (temp1 == false)
            {
                comboBox1.Enabled = false;
                checkBox1.Enabled = false;
                comboBox16.Enabled = false;
                checkBox34.Enabled = false;
                comboBox15.Enabled = false;
                button5.Enabled = false;
            }

            if (r1 == "4")
            {
                string[] temp = { "М8х1" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "3" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "5")
            {
                string[] temp = { "М10х1" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "3" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "6")
            {
                string[] temp = { "М10х1", "М14х1.5", "М12х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "3", "2.5" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "8")
            {
                string[] temp = { "М14х1.5", "М12х1.5", "М16х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "5", "4" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "10")
            {
                string[] temp = { "М16х1.5", "М18х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "7", "6" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "12")
            {
                string[] temp = { "М18х1.5", "М20х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "8" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "15")
            {
                string[] temp = { "М22х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "10" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "16")
            {
                string[] temp = { "М24х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "10", "11" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "18")
            {
                string[] temp = { "М26х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "13" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "22")
            {
                string[] temp = { "М30х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "17" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "28")
            {
                string[] temp = { "М36х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "23" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "34")
            {
                string[] temp = { "М45х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "35")
            {
                string[] temp = { "М45х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "42")
            {
                string[] temp = { "М52х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "36" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "14")
            {
                string[] temp = { "М22х1.5" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "9" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "20")
            {
                string[] temp = { "М30х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "14" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "25")
            {
                string[] temp = { "М36х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "19" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "30")
            {
                string[] temp = { "М42х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "24" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }
            else if (r1 == "38")
            {
                string[] temp = { "М52х2" };
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(temp);

                string[] temp2 = { "32" };
                comboBox16.Items.Clear();
                comboBox16.Items.AddRange(temp2);
            }

        }

        private void comboBox16_TextChanged(object sender, EventArgs e)
        {
            string dv = comboBox16.Text;
            string dn = comboBox10.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox16.Items.Count; i++)
            {
                if (Convert.ToString(dv) == Convert.ToString(comboBox16.Items[i]))
                {
                    if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
                    {
                        string[] temp = { "М10х1", "М12х1.5" };
                        comboBox1.Enabled = true;
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(temp);

                        checkBox1.Enabled = true;
                    }
                    else if (dn == "8" && dv == "5")
                    {
                        string[] temp = { "М14х1.5", "М12х1.5" };
                        comboBox1.Enabled = true;
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(temp);

                        checkBox1.Enabled = true;
                    }
                    else
                    {
                        checkBox15.Enabled = true;
                        temp1 = true;
                        break;
                    }
                }
            }
            if (temp1 == false)
            {
                checkBox15.Enabled = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string rez = comboBox1.Text;
            string dn = comboBox10.Text;
            string dv = comboBox16.Text;

            int flagNip = 100;
            int flagGM = 100;

            Dictionary<string, int> isps = new Dictionary<string, int>();

            if (checkBox3.Checked == true)
            {
                isps = chooseIspsSharTr(dn, dv, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
                //MessageBox.Show($"Семейство GM {flagGM} не найдено. NIP {flagNip} Ошибка!");
            }
            else if (checkBox4.Checked == true)
            {
                isps = chooseIspsSharSht(dn, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }

            if (flagNip >= 10)
            {
                nip23RowName = $"ГОСТ 23355-78 Исполнение 1-{flagNip}";
            }
            else
            {
                nip23RowName = $"ГОСТ 23355-78 Исполнение 1-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-{flagGM}";
            }
            else
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-0{flagGM}";
            }
            Console.WriteLine(g23RowName);
            Console.WriteLine(nip23RowName);

            //Ишу нужную категорию
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, categNameGM);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, categNameNip23);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория {categNameNip23} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g23);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, nip23);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g23} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {nip23} не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();
                        MemberManagerErrorsEnum fr;
                        string fm;

                        //string adressNip = "D:/Users/CADExp/Desktop/ДИПЛОМ_ВКР/Парам дет/ГОСТ 23355-78 Исполнение 1.ipt";
                        //string adressKor = "C:/Users/Coddy/Desktop/Глеб/Программная часть/Итог детали/Соединение с шаровым ниппелем под сварку/ГОСТ 22525-77 Концы корпусных деталей.ipt";
                        //string adressKor = "D:/Users/CADexp/Desktop/Итог детали/Соединение с шаровым ниппелем/ГОСТ 22525-77 Концы корпусных деталей.ipt";

                        //ThisApplication.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, adressKor);
                        //ComponentOccurrence ModelKor = oAssDoc.ComponentDefinition.Occurrences.Add(adressKor, oPositionMatrix);
                        //Inventor.ControlDefinition ctrlDeff2 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff2.Execute();

                        //Добавление из библиотеки компонентов

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow nipTableRow = GetTablelRow(rowsNip, "элемент", nip23RowName);
                        var FileNip = famNip.CreateMember(nipTableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip, oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g23RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);

                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        //Собираем все в одно
                        //Поиск Faces
                        Face oFaceKor1, oFaceNip2, oFaceNip1, oFaceGM1, oFaceNip2tr;

                        oFaceGM1 = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKor1 = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNip1 = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNip2 = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNip2tr = ModelGM.SurfaceBodies[1].Faces[1];

                        Edge test;


                        //Гайка кас/совм {8EF6084F-251D-2041-F149-6900BC6D3FD8} - прошлый
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{D6C563AD-11A3-CAA5-4F14-A4749FC2814A}")
                            {
                                oFaceGM1 = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Корпус
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{57B35F40-4A68-C57F-FEB3-826B8BB7243D}")
                        //    {
                        //        oFaceKor1 = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}
                        //Ниппель вверх кас {E8F736A6-D6E5-CD74-3D7A-ACE21B6D8C1B} - прошлый
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{792A5546-327F-DB7E-D365-DDF7D83AD8FC}")
                            {
                                oFaceNip1 = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель низ вставка 
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{4F92BB6E-EE47-3956-5242-CBE2D8C5BF87}")
                            {
                                oFaceNip2 = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }

                        //Ниппель для трубы
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{52F2D01B-3515-14EF-3251-541626A6B239}")
                            {
                                oFaceNip2tr = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }

                        if (checkBox3.Checked == true)
                        {
                            //Зависимости
                            MateConstraint mate11;
                            //mate11 = oAssCompDef.Constraints.AddMateConstraint(oFaceGM1, oFaceNip1, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);

                            mate11 = oAssCompDef.Constraints.AddMateConstraint(oFaceGM1, oFaceNip1, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            //TangentConstraint tan11;
                            //tan11 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGM1, oFaceNip1, false, 0);

                            oAssCompDef.Constraints.AddInsertConstraint(oFaceGM1, oFaceNip1, false, 0);

                            InsertConstraint insert11;
                            insert11 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNip2tr, faceTr1, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }

                        else if (checkBox4.Checked == true)
                        {

                            //Зависимости
                            MateConstraint mate1;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceGM1, oFaceNip1, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            //TangentConstraint tan1;
                            //tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGM1, oFaceNip1, false, 0);

                            oAssCompDef.Constraints.AddInsertConstraint(oFaceGM1, oFaceNip1, false, 0);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNip2, faceSh1Vs, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }

                        oAssemblyDocName = oAssDoc;
                    }
                }
            }
        }

        //Соединение с коническим ниппелем 1

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            string rez = comboBox2.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox2.Items.Count; i++)
            {
                if (Convert.ToString(rez) == Convert.ToString(comboBox2.Items[i]))
                {
                    checkBox22.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                checkBox22.Enabled = false;
            }
        }
        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            string rez;
            string r1;
            rez = comboBox2.Text;
            r1 = comboBox3.Text;

            bool temp1 = false;
            for (int i = 0; i < comboBox3.Items.Count; i++)
            {
                if (Convert.ToString(r1) == Convert.ToString(comboBox3.Items[i]))
                {
                    if (checkBox7.Checked == true)
                    {
                        comboBox12.Enabled = true;
                        checkBox19.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    else if (checkBox8.Checked == true)
                    {
                        comboBox2.Enabled = true;
                        checkBox21.Enabled = true;
                        temp1 = true;
                        break;
                    }

                }
            }
            if (temp1 == false)
            {
                comboBox12.Enabled = false;
                checkBox19.Enabled = false;
                comboBox2.Enabled = false;
                checkBox21.Enabled = false;
                checkBox22.Enabled = false;
                button7.Enabled = false;
            }

           
            if (r1 == "6")
            {
                string[] temp = { "М10х1", "М14х1.5", "М12х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "3", "2.5" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "8")
            {
                string[] temp = { "М14х1.5", "М12х1.5", "М16х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "5", "4" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "10")
            {
                string[] temp = { "М16х1.5", "М18х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "7", "6" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "12")
            {
                string[] temp = { "М18х1.5", "М20х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "8" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "15")
            {
                string[] temp = { "М22х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "10" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "16")
            {
                string[] temp = { "М24х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "10", "11" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "18")
            {
                string[] temp = { "М26х1.5" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "13" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "22")
            {
                string[] temp = { "М30х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "17" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "28")
            {
                string[] temp = { "М36х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "23" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "35")
            {
                string[] temp = { "М45х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "42")
            {
                string[] temp = { "М52х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "36" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "20")
            {
                string[] temp = { "М30х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "14" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "25")
            {
                string[] temp = { "М36х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "19" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "30")
            {
                string[] temp = { "М42х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "24" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }
            else if (r1 == "38")
            {
                string[] temp = { "М52х2" };
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(temp);

                string[] temp2 = { "32" };
                comboBox12.Items.Clear();
                comboBox12.Items.AddRange(temp2);
            }

        }

        private void comboBox12_TextChanged(object sender, EventArgs e)
        {
            string dv = comboBox12.Text;
            string dn = comboBox3.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox12.Items.Count; i++)
            {
                if (Convert.ToString(dv) == Convert.ToString(comboBox12.Items[i]))
                {
                    if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
                    {
                        string[] temp = { "М10х1", "М12х1.5" };
                        comboBox2.Enabled = true;
                        comboBox2.Items.Clear();
                        comboBox2.Items.AddRange(temp);

                        checkBox21.Enabled = true;
                    }
                    else if (dn == "8" && dv == "5")
                    {
                        string[] temp = { "М14х1.5", "М12х1.5" };
                        comboBox2.Enabled = true;
                        comboBox2.Items.Clear();
                        comboBox2.Items.AddRange(temp);

                        checkBox21.Enabled = true;
                    }
                    else
                    {
                        checkBox22.Enabled = true;
                        temp1 = true;
                        break;
                    }
                }
            }
            if (temp1 == false)
            {
                checkBox22.Enabled = false;
            }



        }

        //Поиск исполнений функции
        public static Dictionary<string, int> chooseIspsKonTr(string dn, string dv, string rez)
        {
            int ispNip = 1;
            int ispGM = 1;
            string dy = "2.5";

            //Выбрал исполнение ниппеля и нашел Dy 
            if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
            {
                ispNip = 1;
                dy = "4";
            }
            else if (dn == "8" && dv == "5")
            {
                ispNip = 2;
                dy = "6";
            }
            else if (dn == "10" && dv == "7")
            {
                ispNip = 3;
                dy = "8";
            }
            else if (dn == "12" && dv == "8")
            {
                ispNip = 4;
                dy = "10";
            }
            else if (dn == "15" && dv == "10")
            {
                ispNip = 5;
                dy = "12";
            }
            else if (dn == "18" && dv == "13")
            {
                ispNip = 6;
                dy = "15";
            }
            else if (dn == "22" && dv == "17")
            {
                ispNip = 7;
                dy = "20";
            }
            else if (dn == "28" && dv == "23")
            {
                ispNip = 8;
                dy = "25";
            }
            else if (dn == "35" && dv == "29")
            {
                ispNip = 9;
                dy = "32";
            }
            else if (dn == "42" && dv == "36")
            {
                ispNip = 10;
                dy = "40";
            }
            else if (dn == "6" && dv == "2.5")
            {
                ispNip = 11;
                dy = "3";
            }
            else if (dn == "8" && dv == "4")
            {
                ispNip = 12;
                dy = "4";
            }
            else if (dn == "10" && dv == "6")
            {
                ispNip = 13;
                dy = "5";
            }
            else if (dn == "12" && dv == "8")
            {
                ispNip = 14;
                dy = "6";
            }
            else if (dn == "16" && dv == "11")
            {
                ispNip = 15;
                dy = "10";
            }
            else if (dn == "20" && dv == "14")
            {
                ispNip = 16;
                dy = "12";
            }
            else if (dn == "25" && dv == "19")
            {
                ispNip = 17;
                dy = "15";
            }
            else if (dn == "30" && dv == "24")
            {
                ispNip = 18;
                dy = "20";
            }
            else if (dn == "38" && dv == "32")
            {
                ispNip = 19;
                dy = "25";
            }

            //Выбираю исполнение гайки по Dn и Dy; в двух местах добавляю rez
            if (dn == "6" && dy == "4")
            {
                if (rez == "М10х1")
                {
                    ispGM = 3;
                }
                else if (rez == "М12х1.5")
                {
                    ispGM = 5;
                }
            }
            else if (dn == "8" && dy == "6")
            {
                if (rez == "М12х1.5")
                {
                    ispGM = 4;
                }
                else if (rez == "М14х1.5")
                {
                    ispGM = 6;
                }
            }
            else if (dn == "10" && dy == "8")
            {
                ispGM = 7;
            }
            else if (dn == "12" && dy == "10")
            {
                ispGM = 8;
            }
            else if (dn == "15" && dy == "12")
            {
                ispGM = 9;
            }
            else if (dn == "18" && dy == "15")
            {
                ispGM = 11;
            }
            else if (dn == "22" && dy == "20")
            {
                ispGM = 12;
            }
            else if (dn == "28" && dy == "25")
            {
                ispGM = 13;
            }
            else if (dn == "35" && dy == "32")
            {
                ispGM = 15;
            }
            else if (dn == "42" && dy == "40")
            {
                ispGM = 16;
            }
            else if (dn == "6" && dy == "3")
            {
                ispGM = 17;
            }
            else if (dn == "8" && dy == "4")
            {
                ispGM = 18;
            }
            else if (dn == "10" && dy == "5")
            {
                ispGM = 19;
            }
            else if (dn == "12" && dy == "6")
            {
                ispGM = 20;
            }
            else if (dn == "16" && dy == "10")
            {
                ispGM = 22;
            }
            else if (dn == "20" && dy == "12")
            {
                ispGM = 23;
            }
            else if (dn == "25" && dy == "15")
            {
                ispGM = 24;
            }
            else if (dn == "30" && dy == "20")
            {
                ispGM = 25;
            }
            else if (dn == "38" && dy == "25")
            {
                ispGM = 26;
            }

            Dictionary<string, int> results = new Dictionary<string, int>();
            results.Add("Nip", ispNip);
            results.Add("Gm", ispGM);

            return results;

        }

        public static Dictionary<string, int> chooseIspsKonSht(string dn, string rez)
        {
            int flagGM = 1;
            int ispNip = 1;
            string dy = "2.5";
            if (dn == "6" && rez == "М10х1")
            {
                flagGM = 3;
                dy = "4";
            }
            else if (dn == "8" && rez == "М12х1.5")
            {
                flagGM = 4;
                dy = "6";
            }
            else if (dn == "6" && rez == "М12х1.5")
            {
                flagGM = 5;
                dy = "4";
            }
            else if (dn == "8" && rez == "М14х1.5")
            {
                flagGM = 6;
                dy = "6";
            }
            else if (dn == "10" && rez == "М16х1.5")
            {
                flagGM = 7;
                dy = "8";
            }
            else if (dn == "12" && rez == "М18х1.5")
            {
                flagGM = 8;
                dy = "10";
            }
            else if (dn == "15" && rez == "М22х1.5")
            {
                flagGM = 9;
                dy = "12";
            }
            else if (dn == "18" && rez == "М26х1.5")
            {
                flagGM = 11;
                dy = "15";
            }
            else if (dn == "22" && rez == "М30х2")
            {
                flagGM = 12;
                dy = "20";
            }
            else if (dn == "28" && rez == "М36х2")
            {
                flagGM = 13;
                dy = "25";
            }
            else if (dn == "35" && rez == "М45х2")
            {
                flagGM = 15;
                dy = "32";
            }
            else if (dn == "42" && rez == "М52х2")
            {
                flagGM = 16;
                dy = "40";
            }
            //Group 3
            else if (dn == "6" && rez == "М14х1.5")
            {
                flagGM = 17;
                dy = "3";
            }
            else if (dn == "8" && rez == "М16х1.5")
            {
                flagGM = 18;
                dy = "4";
            }
            else if (dn == "10" && rez == "М18х1.5")
            {
                flagGM = 19;
                dy = "5";
            }
            else if (dn == "12" && rez == "М20х1.5")
            {
                flagGM = 20;
                dy = "6";
            }
            else if (dn == "16" && rez == "М24х1.5")
            {
                flagGM = 22;
                dy = "10";
            }
            else if (dn == "20" && rez == "М30х2")
            {
                flagGM = 23;
                dy = "12";
            }
            else if (dn == "25" && rez == "М36х2")
            {
                flagGM = 24;
                dy = "15";
            }
            else if (dn == "30" && rez == "М42х2")
            {
                flagGM = 25;
                dy = "20";
            }
            else if (dn == "38" && rez == "М52х2")
            {
                flagGM = 26;
                dy = "25";
            }

            //Поиск исполнения ниппеля по Dn и Dy
            if (dn == "6" && dy == "4") // Открываем резьбу в комбо боксах
            {
                ispNip = 1;
            }
            else if (dn == "8" && dy == "6")
            {
                ispNip = 2;
            }
            else if (dn == "10" && dy == "8")
            {
                ispNip = 3;
            }
            else if (dn == "12" && dy == "10")
            {
                ispNip = 4;
            }
            else if (dn == "15" && dy == "12")
            {
                ispNip = 5;
            }
            else if (dn == "18" && dy == "15")
            {
                ispNip = 6;
            }
            else if (dn == "22" && dy == "20")
            {
                ispNip = 7;
            }
            else if (dn == "28" && dy == "25")
            {
                ispNip = 8;
            }
            else if (dn == "35" && dy == "32")
            {
                ispNip = 9;
            }
            else if (dn == "42" && dy == "40")
            {
                ispNip = 10;
            }
            else if (dn == "6" && dy == "3")
            {
                ispNip = 11;
            }
            else if (dn == "8" && dy == "4")
            {
                ispNip = 12;
            }
            else if (dn == "10" && dy == "5")
            {
                ispNip = 13;
            }
            else if (dn == "12" && dy == "6")
            {
                ispNip = 14;
            }
            else if (dn == "16" && dy == "10")
            {
                ispNip = 15;
            }
            else if (dn == "20" && dy == "12")
            {
                ispNip = 16;
            }
            else if (dn == "25" && dy == "15")
            {
                ispNip = 17;
            }
            else if (dn == "30" && dy == "20")
            {
                ispNip = 18;
            }
            else if (dn == "38" && dy == "25")
            {
                ispNip = 19;
            }

            Dictionary<string, int> results = new Dictionary<string, int>();
            results.Add("Nip", ispNip);
            results.Add("Gm", flagGM);

            return results;

        }


        private void button7_Click(object sender, EventArgs e)
        {
            string rez = comboBox2.Text;
            string dv = comboBox12.Text;
            string dn = comboBox3.Text;

            int flagNip = 100;
            int flagGM = 100;

            Dictionary<string, int> isps = new Dictionary<string, int>();

            if (checkBox7.Checked == true)
            {
                isps = chooseIspsKonTr(dn, dv, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }
            else if (checkBox8.Checked == true)
            {
                isps = chooseIspsKonSht(dn, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }

            if (flagNip >= 10)
            {
                konNip1RowName = $"ГОСТ 28016-89 Исполнение 1-{flagNip}";
            }
            else
            {
                konNip1RowName = $"ГОСТ 28016-89 Исполнение 1-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-{flagGM}";
            }
            else
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-0{flagGM}";
            }
            Console.WriteLine(g23RowName);
            Console.WriteLine(nip23RowName);

            //Ишу нужную категорию
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, categNameGM);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, categNameKonNip1);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория {categNameKonNip1} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g23);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, konNip1);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g23} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {konNip1} не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();
                        MemberManagerErrorsEnum fr;
                        string fm;

                        //string adressNip = "D:/Users/CADExp/Desktop/ДИПЛОМ_ВКР/Парам дет/ГОСТ 23355-78 Исполнение 1.ipt";
                        //string adressKor = "C:/Users/Coddy/Desktop/Глеб/Программная часть/Итог детали/Соединение с коническим ниппелем/ГОСТ 22525-77 Концы корпусных деталей с углом конуса 24°.ipt";
                        //string adressKor = "D:/Users/CADexp/Desktop/Итог детали/Соединение с шаровым ниппелем/ГОСТ 22525-77 Концы корпусных деталей.ipt";

                        //ThisApplication.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, adressKor);
                        //ComponentOccurrence ModelKor = oAssDoc.ComponentDefinition.Occurrences.Add(adressKor, oPositionMatrix);
                        //Inventor.ControlDefinition ctrlDeff2 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff2.Execute();

                        //Добавление из библиотеки компонентов

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow konNip1TableRow = GetTablelRow(rowsNip, "элемент", konNip1RowName);
                        var FileNip = famNip.CreateMember(konNip1TableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip, oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g23RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);

                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        //Прокси оси

                        //Для ниппеля
                        Inventor.PartComponentDefinition pPartDef10 = null;
                        pPartDef10 = ModelNip.Definition as PartComponentDefinition;
                        //получение рабочей оси
                        Inventor.WorkAxis pPartAxis10 = pPartDef10.WorkAxes[3];
                        //преобразование оси в прокси геометрии сборки
                        object axisProxy10 = null;
                        Inventor.WorkAxisProxy pWorkAxisProxyKorp;
                        ModelNip.CreateGeometryProxy(pPartAxis10, out axisProxy10);
                        pWorkAxisProxyKorp = axisProxy10 as WorkAxisProxy;

                        //Для гайки
                        Inventor.PartComponentDefinition pPartDef11 = null;
                        pPartDef11 = ModelGM.Definition as PartComponentDefinition;
                        //получение рабочей оси
                        Inventor.WorkAxis pPartAxis11 = pPartDef11.WorkAxes[3];
                        //преобразование оси в прокси геометрии сборки
                        object axisProxy11 = null;
                        Inventor.WorkAxisProxy pWorkAxisProxyKorp1;
                        ModelGM.CreateGeometryProxy(pPartAxis11, out axisProxy11);
                        pWorkAxisProxyKorp1 = axisProxy11 as WorkAxisProxy;


                        //Собираем все в одно

                        Face oFaceGMOs, oFaceGMKas, oFaceNipOs, oFaceNipVs, oFaceNipKas, oFaceKorOs, oFaceKorVs, oFaceKorTr, oFaceGMVs, oFaceNipVsGm;

                        oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorTr = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipVsGm = ModelGM.SurfaceBodies[1].Faces[1];

                        Face GMtemp = ModelGM.SurfaceBodies[1].Faces[1];
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{8EF6084F-251D-2041-F149-6900BC6D3FD8}")
                            {
                                GMtemp = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        Face NPtemp = ModelGM.SurfaceBodies[1].Faces[1];
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{7CF737AC-B282-926F-2D47-67B0EF1C05A4}")
                            {
                                NPtemp = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }

                        //Гайка ось
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{563551C3-F11F-C251-1BE8-C9CF1BF4E066}")
                            {
                                oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка вставка к ниппелю
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{D6C563AD-11A3-CAA5-4F14-A4749FC2814A}")
                            {
                                oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка кас
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{8EF6084F-251D-2041-F149-6900BC6D3FD8}")
                            {
                                oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель ось
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{29347A13-1984-3A98-1C95-835915682BC3}")
                            {
                                oFaceNipOs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель вставка к гайке
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{29347A13-1984-3A98-1C95-835915682BC3}")
                            {
                                oFaceNipVsGm = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель кас
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{FB243FD5-813F-3C44-D216-E79E6C6BE86D}")
                            {
                                oFaceNipKas = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель вставка
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{6767762D-5FC3-17F0-47B6-7E540FBB4388}")
                            {
                                oFaceNipVs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }

                        //Ниппель для трубы
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{53A91A09-9AB3-FA79-D00A-91562FBE6BDD}")
                            {
                                oFaceKorTr = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Корпус ось {29347A13-1984-3A98-1C95-835915682BC3}
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{CAE4D3B7-A292-0A03-12F7-B3AE986B98F7}")
                        //    {
                        //        oFaceKorOs = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}
                        ////Корпус вставка
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{714B8F73-E39D-FD75-B70C-F6ED3E688C68}")
                        //    {
                        //        oFaceKorVs = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}

                        if (checkBox8.Checked == true)
                        {
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceGMOs, oFaceNipOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);


                            //Гайка и ниппель вместе
                            //InsertConstraint insert2;
                            //insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipVsGm, GMtemp, false, -0.2);
                            InsertConstraint insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMKas, oFaceNipVsGm, true, 0);

                            //TangentConstraint tan2 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGMKas, NPtemp, true, 0);
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipKas, faceKon1Vs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceNipKas, faceKon1Vs, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }
                        else if (checkBox7.Checked == true)
                        {
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(GMtemp, NPtemp, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            //mate2 = oAssCompDef.Constraints.AddMateConstraint(faceKon1Os, oFaceGMOs, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);

                            //TangentConstraint tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGMKas, NPtemp, true, 0);
                            InsertConstraint insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMKas, oFaceNipVsGm, true, 0);
                            //Гайка и ниппель вместе
                            //InsertConstraint insert2;
                            //insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipVsGm, GMtemp, false, -0.2);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceKorTr, faceTr3, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }


                        oAssemblyDocName = oAssDoc;


                    }
                }
            }
        }

        //Соединение с шаровым ниппелем на пайке

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            
            string rez = comboBox4.Text;
            //comboBox1.Update();
            //comboBox10.Update();
            bool temp1 = false;
            for (int i = 0; i < comboBox4.Items.Count; i++)
            {
                if (Convert.ToString(rez) == Convert.ToString(comboBox4.Items[i]))
                {
                    checkBox18.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                checkBox18.Enabled = false;
            }
        }


        //Проверка изменения наружного диаметра
        private void comboBox11_TextChanged(object sender, EventArgs e)
        {
            string rez;
            string r1;
            rez = comboBox4.Text;
            r1 = comboBox11.Text;

            bool temp1 = false;
            for (int i = 0; i < comboBox11.Items.Count; i++)
            {
                if (Convert.ToString(r1) == Convert.ToString(comboBox11.Items[i]))
                {
                    if (checkBox6.Checked == true)
                    {
                        comboBox4.Enabled = true;
                        checkBox17.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    else if (checkBox5.Checked == true)
                    {
                        comboBox15.Enabled = true;
                        checkBox33.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    
                }
            }
            if (temp1 == false)
            {
                comboBox4.Enabled = false;
                checkBox17.Enabled = false;
                comboBox15.Enabled = false;
                checkBox33.Enabled = false;
                checkBox18.Enabled = false;
                button8.Enabled = false;
            }

            if (r1 == "4")
            {
                string[] temp = { "М8х1" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "3" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "5")
            {
                string[] temp = { "М10х1" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "3" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "6")
            {
                string[] temp = { "М10х1", "М14х1.5", "М12х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "3", "2.5" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "8")
            {
                string[] temp = { "М14х1.5", "М12х1.5", "М16х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "5", "4" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "10")
            {
                string[] temp = { "М16х1.5", "М18х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "7", "6" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "12")
            {
                string[] temp = { "М18х1.5", "М20х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "8" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "15")
            {
                string[] temp = { "М22х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "10" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "16")
            {
                string[] temp = { "М24х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "10", "11" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "18")
            {
                string[] temp = { "М26х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "13" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "22")
            {
                string[] temp = { "М30х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "17" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "28")
            {
                string[] temp = { "М36х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "23" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "34")
            {
                string[] temp = { "М45х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "35")
            {
                string[] temp = { "М45х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "42")
            {
                string[] temp = { "М52х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "36" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "14")
            {
                string[] temp = { "М22х1.5" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "9" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "20")
            {
                string[] temp = { "М30х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "14" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "25")
            {
                string[] temp = { "М36х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "19" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "30")
            {
                string[] temp = { "М42х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "24" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
            else if (r1 == "38")
            {
                string[] temp = { "М52х2" };
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(temp);

                string[] temp2 = { "32" };
                comboBox15.Items.Clear();
                comboBox15.Items.AddRange(temp2);
            }
        }

        //Внутренний диаметр
        private void comboBox15_TextChanged(object sender, EventArgs e)
        {
            string dv = comboBox15.Text;
            string dn = comboBox11.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox15.Items.Count; i++)
            {
                if (Convert.ToString(dv) == Convert.ToString(comboBox15.Items[i]))
                {
                    if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
                    {
                        string[] temp = { "М10х1", "М12х1.5" };
                        comboBox4.Enabled = true;
                        comboBox4.Items.Clear();
                        comboBox4.Items.AddRange(temp);
                        
                        checkBox17.Enabled = true;
                    }
                    else if (dn == "8" && dv == "5")
                    {
                        string[] temp = { "М14х1.5", "М12х1.5" };
                        comboBox4.Enabled = true;
                        comboBox4.Items.Clear();
                        comboBox4.Items.AddRange(temp);
                        
                        checkBox17.Enabled = true;
                    }
                    else
                    {
                        checkBox18.Enabled = true;
                        temp1 = true;
                        break;
                    }
                }
            }
            if (temp1 == false)
            {
                checkBox18.Enabled = false;
            }
        }

        // Шаровой ниппель 1 новый поиск

        // Ниппель шаровой (при трубе у пользователя)
        public static Dictionary<string, int> chooseIspsSharTr(string dn, string dv, string rez)
        {

            int ispNip = 1;
            int ispGM = 1;
            string dy = "2.5";

            //Выбор исполнения ниппеля и Dy 
            if (dn == "4" && dv == "3")
            {
                ispNip = 1;
                dy = "2.5";
            }
            else if (dn == "5" && dv == "3")
            {
                ispNip = 2;
                dy = "3";
            }
            else if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
            {
                ispNip = 3;
                dy = "4";
            }
            else if (dn == "8" && dv == "5")
            {
                ispNip = 4;
                dy = "6";
            }
            else if (dn == "10" && dv == "7")
            {
                ispNip = 5;
                dy = "8";
            }
            else if (dn == "12" && dv == "8")
            {
                ispNip = 6;
                dy = "10";
            }
            else if (dn == "15" && dv == "10")
            {
                ispNip = 7;
                dy = "12";
            }
            else if (dn == "16" && dv == "10")
            {
                ispNip = 8;
                dy = "12";
            }
            else if (dn == "18" && dv == "13")
            {
                ispNip = 9;
                dy = "15";
            }
            else if (dn == "22" && dv == "17")
            {
                ispNip = 10;
                dy = "20";
            }
            else if (dn == "28" && dv == "23")
            {
                ispNip = 11;
                dy = "25";
            }
            else if (dn == "34" && dv == "29")
            {
                ispNip = 12;
                dy = "32";
            }
            else if (dn == "35" && dv == "29")
            {
                ispNip = 13;
                dy = "32";
            }
            else if (dn == "42" && dv == "36")
            {
                ispNip = 14;
                dy = "40";
            }
            else if (dn == "6" && dv == "2.5")
            {
                ispNip = 15;
                dy = "3";
            }
            else if (dn == "8" && dv == "4")
            {
                ispNip = 16;
                dy = "4";
            }
            else if (dn == "10" && dv == "6")
            {
                ispNip = 17;
                dy = "5";
            }
            else if (dn == "12" && dv == "8")
            {
                ispNip = 18;
                dy = "6";
            }
            else if (dn == "14" && dv == "9")
            {
                ispNip = 19;
                dy = "8";
            }
            else if (dn == "16" && dv == "11")
            {
                ispNip = 20;
                dy = "10";
            }
            else if (dn == "20" && dv == "14")
            {
                ispNip = 21;
                dy = "12";
            }
            else if (dn == "25" && dv == "19")
            {
                ispNip = 22;
                dy = "15";
            }
            else if (dn == "30" && dv == "24")
            {
                ispNip = 23;
                dy = "20";
            }
            else if (dn == "38" && dv == "32")
            {
                ispNip = 24;
                dy = "25";
            }

            //Выбираю исполнение гайки по Dn и Dy; в двух местах добавляю rez
            if (dn == "4" && dy == "2.5")
            {
                ispGM = 1;
            }
            else if (dn == "5" && dy == "3")
            {
                ispGM = 2;
            }
            else if (dn == "6" && dy == "4")
            {
                if (rez == "М10х1")
                {
                    ispGM = 3;
                }
                else if (rez == "М12х1.5")
                {
                    ispGM = 5;
                }
            }
            else if (dn == "8" && dy == "6")
            {
                if (rez == "М12х1.5")
                {
                    ispGM = 4;
                }
                else if (rez == "М14х1.5")
                {
                    ispGM = 6;
                }
            }
            else if (dn == "10" && dy == "8")
            {
                ispGM = 7;
            }
            else if (dn == "12" && dy == "10")
            {
                ispGM = 8;
            }
            else if (dn == "15" && dy == "12")
            {
                ispGM = 9;
            }
            else if (dn == "16" && dy == "12")
            {
                ispGM = 10;
            }
            else if (dn == "18" && dy == "15")
            {
                ispGM = 11;
            }
            else if (dn == "22" && dy == "20")
            {
                ispGM = 12;
            }
            else if (dn == "28" && dy == "25")
            {
                ispGM = 13;
            }
            else if (dn == "34" && dy == "32")
            {
                ispGM = 14;
            }
            else if (dn == "35" && dy == "32")
            {
                ispGM = 15;
            }
            else if (dn == "42" && dy == "40")
            {
                ispGM = 16;
            }
            else if (dn == "6" && dy == "3")
            {
                ispGM = 17;
            }
            else if (dn == "8" && dy == "4")
            {
                ispGM = 18;
            }
            else if (dn == "10" && dy == "5")
            {
                ispGM = 19;
            }
            else if (dn == "12" && dy == "6")
            {
                ispGM = 20;
            }
            else if (dn == "14" && dy == "8")
            {
                ispGM = 21;
            }
            else if (dn == "16" && dy == "10")
            {
                ispGM = 22;
            }
            else if (dn == "20" && dy == "12")
            {
                ispGM = 23;
            }
            else if (dn == "25" && dy == "15")
            {
                ispGM = 24;
            }
            else if (dn == "30" && dy == "20")
            {
                ispGM = 25;
            }
            else if (dn == "38" && dy == "25")
            {
                ispGM = 26;
            }

            Dictionary<string, int> results = new Dictionary<string, int>();
            results.Add("Nip", ispNip);
            results.Add("Gm", ispGM);

            return results;

        }

        // Ниппель шаровой (при штуцере у пользователя)
        public static Dictionary<string, int> chooseIspsSharSht(string dn, string rez)
        {

            int flagGM = 1;
            int ispNip = 1;
            string dy = "2.5";
            if (dn == "4" && rez == "М8х1")
            {
                flagGM = 1;
                dy = "2.5";
            }
            else if (dn == "5" && rez == "М10х1")
            {
                flagGM = 2;
                dy = "3";
            }
            else if (dn == "6" && rez == "М10х1")
            {
                flagGM = 3;
                dy = "4";
            }
            else if (dn == "8" && rez == "М12х1.5")
            {
                flagGM = 4;
                dy = "6";
            }
            else if (dn == "6" && rez == "М12х1.5")
            {
                flagGM = 5;
                dy = "4";
            }
            else if (dn == "8" && rez == "М14х1.5")
            {
                flagGM = 6;
                dy = "6";
            }
            else if (dn == "10" && rez == "М16х1.5")
            {
                flagGM = 7;
                dy = "8";
            }
            else if (dn == "12" && rez == "М18х1.5")
            {
                flagGM = 8;
                dy = "10";
            }
            else if (dn == "15" && rez == "М22х1.5")
            {
                flagGM = 9;
                dy = "12";
            }
            else if (dn == "16" && rez == "М24х1.5")
            {
                flagGM = 10;
                dy = "12";
            }
            else if (dn == "18" && rez == "М26х1.5")
            {
                flagGM = 11;
                dy = "15";
            }
            else if (dn == "22" && rez == "М30х2")
            {
                flagGM = 12;
                dy = "20";
            }
            else if (dn == "28" && rez == "М36х2")
            {
                flagGM = 13;
                dy = "25";
            }
            else if (dn == "34" && rez == "М45х2")
            {
                flagGM = 14;
                dy = "32";
            }
            else if (dn == "35" && rez == "М45х2")
            {
                flagGM = 15;
                dy = "32";
            }
            else if (dn == "42" && rez == "М52х2")
            {
                flagGM = 16;
                dy = "40";
            }
            //Group 3
            else if (dn == "6" && rez == "М14х1.5")
            {
                flagGM = 17;
                dy = "3";
            }
            else if (dn == "8" && rez == "М16х1.5")
            {
                flagGM = 18;
                dy = "4";
            }
            else if (dn == "10" && rez == "М18х1.5")
            {
                flagGM = 19;
                dy = "5";
            }
            else if (dn == "12" && rez == "М20х1.5")
            {
                flagGM = 20;
                dy = "6";
            }
            else if (dn == "14" && rez == "М22х1.5")
            {
                flagGM = 21;
                dy = "8";
            }
            else if (dn == "16" && rez == "М24х1.5")
            {
                flagGM = 22;
                dy = "10";
            }
            else if (dn == "20" && rez == "М30х2")
            {
                flagGM = 23;
                dy = "12";
            }
            else if (dn == "25" && rez == "М36х2")
            {
                flagGM = 24;
                dy = "15";
            }
            else if (dn == "30" && rez == "М42х2")
            {
                flagGM = 25;
                dy = "20";
            }
            else if (dn == "38" && rez == "М52х2")
            {
                flagGM = 26;
                dy = "25";
            }

            //Поиск исполнения ниппеля по Dn и Dy
            if (dn == "4" && dy == "2.5")
            {
                ispNip = 1;
            }
            else if (dn == "5" && dy == "3")
            {
                ispNip = 2;
            }
            else if (dn == "6" && dy == "4") // Открываем резьбу в комбо боксах
            {
                ispNip = 3;
            }
            else if (dn == "8" && dy == "6")
            {
                ispNip = 4;
            }
            else if (dn == "10" && dy == "8")
            {
                ispNip = 5;
            }
            else if (dn == "12" && dy == "10")
            {
                ispNip = 6;
            }
            else if (dn == "15" && dy == "12")
            {
                ispNip = 7;
            }
            else if (dn == "16" && dy == "12")
            {
                ispNip = 8;
            }
            else if (dn == "18" && dy == "15")
            {
                ispNip = 9;
            }
            else if (dn == "22" && dy == "20")
            {
                ispNip = 10;
            }
            else if (dn == "28" && dy == "25")
            {
                ispNip = 11;
            }
            else if (dn == "34" && dy == "32")
            {
                ispNip = 12;
            }
            else if (dn == "35" && dy == "32")
            {
                ispNip = 13;
            }
            else if (dn == "42" && dy == "40")
            {
                ispNip = 14;
            }
            else if (dn == "6" && dy == "3")
            {
                ispNip = 15;
            }
            else if (dn == "8" && dy == "4")
            {
                ispNip = 16;
            }
            else if (dn == "10" && dy == "5")
            {
                ispNip = 17;
            }
            else if (dn == "12" && dy == "6")
            {
                ispNip = 18;
            }
            else if (dn == "14" && dy == "8")
            {
                ispNip = 19;
            }
            else if (dn == "16" && dy == "10")
            {
                ispNip = 20;
            }
            else if (dn == "20" && dy == "12")
            {
                ispNip = 21;
            }
            else if (dn == "25" && dy == "15")
            {
                ispNip = 22;
            }
            else if (dn == "30" && dy == "20")
            {
                ispNip = 23;
            }
            else if (dn == "38" && dy == "25")
            {
                ispNip = 24;
            }

            Dictionary<string, int> results = new Dictionary<string, int>();
            results.Add("Nip", ispNip);
            results.Add("Gm", flagGM);

            return results;

        }


        private void button8_Click(object sender, EventArgs e)
        {
            string rez = comboBox4.Text;
            string dn = comboBox11.Text;
            string dv = comboBox15.Text;

            int flagNip = 100;
            int flagGM = 100;

            Dictionary<string, int> isps = new Dictionary<string, int>();

            if (checkBox5.Checked == true)
            {
                isps = chooseIspsSharTr(dn, dv, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }
            else if (checkBox6.Checked == true)
            {
                isps = chooseIspsSharSht(dn, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }
            

            if (flagNip >= 10)
            {
                nipSH2RowName = $"ГОСТ 23355-78 Исполнение 2-{flagNip}";
            }
            else
            {
                nipSH2RowName = $"ГОСТ 23355-78 Исполнение 2-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-{flagGM}";
            }
            else
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-0{flagGM}";
            }

            //Поиск нужных категорий
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, 
                categNameGM);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, 
                categNameNipSH2);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория " +
                $"{categNameNipSH2} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g23);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, nipSH2);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g23} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {nipSH2} " +
                        $"не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();

                        //Добавление из библиотеки компонентов
                        MemberManagerErrorsEnum fr;
                        string fm;

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow nipTableRow = GetTablelRow(rowsNip, "элемент", nipSH2RowName);
                        var FileNip = famNip.CreateMember(nipTableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip,
                            oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g23RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);

                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        //Прокси 

                        //Для ниппеля
                        Inventor.PartComponentDefinition pPartDef10 = null;
                        pPartDef10 = ModelNip.Definition as PartComponentDefinition;
                        //получение рабочей оси
                        Inventor.WorkAxis pPartAxis10 = pPartDef10.WorkAxes[3];
                        //преобразование оси в прокси геометрии сборки
                        object axisProxy10 = null;
                        Inventor.WorkAxisProxy pWorkAxisProxyKorp;
                        ModelNip.CreateGeometryProxy(pPartAxis10, out axisProxy10);
                        pWorkAxisProxyKorp = axisProxy10 as WorkAxisProxy;

                        //Для гайки
                        Inventor.PartComponentDefinition pPartDef11 = null;
                        pPartDef11 = ModelGM.Definition as PartComponentDefinition;
                        //получение рабочей оси
                        Inventor.WorkAxis pPartAxis11 = pPartDef11.WorkAxes[3];
                        //преобразование оси в прокси геометрии сборки
                        object axisProxy11 = null;
                        Inventor.WorkAxisProxy pWorkAxisProxyKorp1;
                        ModelGM.CreateGeometryProxy(pPartAxis11, out axisProxy11);
                        pWorkAxisProxyKorp1 = axisProxy11 as WorkAxisProxy;

                        //Собираем все в одно
                        //Поиск Faces
                        Face oFaceKorNiz, oFaceNipNiz, oFaceNipOs, oFaceNipKas, oFaceGMKas, oFaceGMOs, oFaceNipTr;

                        oFaceKorNiz = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipNiz = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipTr = ModelGM.SurfaceBodies[1].Faces[1];

                        //Ниппель ось
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{405AA019-B1C3-EA85-1CD8-35B9256AFFF6}")
                            {
                                oFaceNipOs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель низ {75D8D5E9-8D7E-B0C8-1F87-672866CF3D29} {F3CA54DB-D5C0-FD52-388F-614F471C149A}
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{F3CA54DB-D5C0-FD52-388F-614F471C149A}")
                            {
                                oFaceNipNiz = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель кас
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{75D8D5E9-8D7E-B0C8-1F87-672866CF3D29}")
                            {
                                oFaceNipKas = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка кас
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{8EF6084F-251D-2041-F149-6900BC6D3FD8}")
                            {
                                oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка ось
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{563551C3-F11F-C251-1BE8-C9CF1BF4E066}")
                            {
                                oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель для трубы
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{F005DDD7-745C-7F69-1199-C2FB644C846A}")
                            {
                                oFaceNipTr = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Корпус низ {F005DDD7-745C-7F69-1199-C2FB644C846A}
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{57B35F40-4A68-C57F-FEB3-826B8BB7243D}")
                        //    {
                        //        oFaceKorNiz = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}

                        if (checkBox6.Checked == true)
                        {
                            MateConstraint mate1;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0,
                                InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceNipKas, oFaceGMKas, true, 0);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipNiz, faceSh2Vs, true, 0);

                            chooseSelection(false);
                        }
                        else if (checkBox5.Checked == true)
                        {
                            MateConstraint mate1;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, 
                                InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceNipKas, oFaceGMKas, true, 0);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipTr, faceTr2, true, 0);

                            chooseSelection(false);
                        }
                        

                        oAssemblyDocName = oAssDoc;

                    }
                }
            }
        }

        //Конический ниппель 2
        private void comboBox5_TextChanged(object sender, EventArgs e)
        {
            string rez = comboBox5.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox5.Items.Count; i++)
            {
                if (Convert.ToString(rez) == Convert.ToString(comboBox5.Items[i]))
                {
                    checkBox26.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                checkBox26.Enabled = false;
            }
        }

        private void comboBox14_TextChanged(object sender, EventArgs e)
        {
            string rez;
            string r1;
            rez = comboBox2.Text;
            r1 = comboBox14.Text;

            bool temp1 = false;
            for (int i = 0; i < comboBox14.Items.Count; i++)
            {
                if (Convert.ToString(r1) == Convert.ToString(comboBox14.Items[i]))
                {
                    if (checkBox9.Checked == true)
                    {
                        comboBox13.Enabled = true;
                        checkBox23.Enabled = true;
                        temp1 = true;
                        break;
                    }
                    else if (checkBox10.Checked == true)
                    {
                        comboBox5.Enabled = true;
                        checkBox25.Enabled = true;
                        temp1 = true;
                        break;
                    }

                }
            }
            if (temp1 == false)
            {
                comboBox13.Enabled = false;
                checkBox23.Enabled = false;
                comboBox5.Enabled = false;
                checkBox25.Enabled = false;
                checkBox26.Enabled = false;
                button10.Enabled = false;
            }


            if (r1 == "6")
            {
                string[] temp = { "М10х1", "М14х1.5", "М12х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "3", "2.5" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "8")
            {
                string[] temp = { "М14х1.5", "М12х1.5", "М16х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "5", "4" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "10")
            {
                string[] temp = { "М16х1.5", "М18х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "7", "6" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "12")
            {
                string[] temp = { "М18х1.5", "М20х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "8" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "15")
            {
                string[] temp = { "М22х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "10" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "16")
            {
                string[] temp = { "М24х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "10", "11" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "18")
            {
                string[] temp = { "М26х1.5" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "13" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "22")
            {
                string[] temp = { "М30х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "17" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "28")
            {
                string[] temp = { "М36х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "23" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "35")
            {
                string[] temp = { "М45х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "29" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "42")
            {
                string[] temp = { "М52х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "36" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "20")
            {
                string[] temp = { "М30х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "14" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "25")
            {
                string[] temp = { "М36х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "19" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "30")
            {
                string[] temp = { "М42х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "24" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
            else if (r1 == "38")
            {
                string[] temp = { "М52х2" };
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(temp);

                string[] temp2 = { "32" };
                comboBox13.Items.Clear();
                comboBox13.Items.AddRange(temp2);
            }
        }

        private void comboBox13_TextChanged(object sender, EventArgs e)
        {

            string dv = comboBox13.Text;
            string dn = comboBox14.Text;
            bool temp1 = false;
            for (int i = 0; i < comboBox13.Items.Count; i++)
            {
                if (Convert.ToString(dv) == Convert.ToString(comboBox13.Items[i]))
                {
                    if (dn == "6" && dv == "3") // Открываем резьбу в комбо боксах
                    {
                        string[] temp = { "М10х1", "М12х1.5" };
                        comboBox5.Enabled = true;
                        comboBox5.Items.Clear();
                        comboBox5.Items.AddRange(temp);

                        checkBox25.Enabled = true;
                    }
                    else if (dn == "8" && dv == "5")
                    {
                        string[] temp = { "М14х1.5", "М12х1.5" };
                        comboBox5.Enabled = true;
                        comboBox5.Items.Clear();
                        comboBox5.Items.AddRange(temp);

                        checkBox25.Enabled = true;
                    }
                    else
                    {
                        checkBox26.Enabled = true;
                        temp1 = true;
                        break;
                    }
                }
            }
            if (temp1 == false)
            {
                checkBox26.Enabled = false;
            }

            //string rez = "";
            //string dy = "";
            //string dn = "";
            //rez = comboBox5.Text;
            //dy = comboBox13.Text;
            //dn = comboBox14.Text;

            //bool temp1 = false;
            //for (int i = 0; i < comboBox13.Items.Count; i++)
            //{
            //    if (Convert.ToString(dy) == Convert.ToString(comboBox13.Items[i]))
            //    {
            //        comboBox14.Enabled = true;
            //        checkBox24.Enabled = true;
            //        temp1 = true;
            //        break;
            //    }
            //}
            //if (temp1 == false)
            //{
            //    comboBox14.Enabled = false;
            //    checkBox24.Enabled = false;
            //}


            //if (dy == "4")
            //{
            //    string[] temp = { "6", "8" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "6")
            //{
            //    string[] temp = { "8", "12" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "8")
            //{
            //    string[] temp = { "10" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "10")
            //{
            //    string[] temp = { "12", "16" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "12")
            //{
            //    string[] temp = { "15", "20" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "15")
            //{
            //    string[] temp = { "18", "25" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "20")
            //{
            //    string[] temp = { "22", "30" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "25")
            //{
            //    string[] temp = { "28", "38" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "32")
            //{
            //    string[] temp = { "35" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "40")
            //{
            //    string[] temp = { "42" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "3")
            //{
            //    string[] temp = { "6" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
            //else if (dy == "5")
            //{
            //    string[] temp = { "10" };
            //    comboBox14.Items.Clear();
            //    comboBox14.Items.AddRange(temp);
            //}
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string rez = comboBox5.Text;
            string dv = comboBox13.Text;
            string dn = comboBox14.Text;

            int flagNip = 100;
            int flagGM = 100;

            Dictionary<string, int> isps = new Dictionary<string, int>();

            if (checkBox9.Checked == true)
            {
                isps = chooseIspsKonTr(dn, dv, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }
            else if (checkBox10.Checked == true)
            {
                isps = chooseIspsKonSht(dn, rez);
                flagNip = isps["Nip"];
                flagGM = isps["Gm"];
            }

            
            if (flagNip == 1)
            {
                ringRowName = "005-007-14n";
            }
            else if (flagNip == 2)
            {
                ringRowName = "008-010-14n";
            }
            else if (flagNip == 3)
            {
                ringRowName = "009-012-19n";
            }//
            else if (flagNip == 4)
            {
                ringRowName = "010-013-19n";
            }
            else if (flagNip == 5)
            {
                ringRowName = "013-016-19n";
            }
            else if (flagNip == 6)
            {
                ringRowName = "016-019-19n";
            }
            else if (flagNip == 7)
            {
                ringRowName = "020-023-19n";
            }
            else if (flagNip == 8)
            {
                ringRowName = "026-029-19n";
            }
            else if (flagNip == 9)
            {
                //flagGM = 15; // в гайке нет этого исполнения добавить!!! уже есть, но все равно проверить
                ringRowName = "032-036-25n";
            }
            else if (flagNip == 10)
            {
                ringRowName = "040-044-25n";
            }
            else if (flagNip == 11)
            {
                ringRowName = "005-007-14n";
            }
            else if (flagNip == 12)
            {
                ringRowName = "008-010-14n";
            }
            else if (flagNip == 13)
            {
                ringRowName = "009-012-19n";
            }
            else if (flagNip == 14)
            {
                ringRowName = "010-013-19n";
            }
            else if (flagNip == 15)
            {
                ringRowName = "013-016-19n";
            }
            else if (flagNip == 16)
            {
                ringRowName = "017-021-25n";
            }
            else if (flagNip == 17)
            {
                ringRowName = "021-025-25n";
            }
            else if (flagNip == 18)
            {
                ringRowName = "026-030-25n";
            }
            else if (flagNip == 19)
            {
                ringRowName = "034-038-25n";
            }

            if (flagNip >= 10)
            {
                konNip2RowName = $"ГОСТ 28016-89 Исполнение_2-{flagNip}";
            }
            else
            {
                konNip2RowName = $"ГОСТ 28016-89 Исполнение_2-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-{flagGM}";
            }
            else
            {
                g23RowName = $"ГОСТ 23353-78 ГАЙКИ НАКИДНЫЕ М8-0{flagGM}";
            }

            //Ишу нужную категорию
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, categNameGM);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, categNameKonNip2);
            ContentTreeViewNode listCategoryNodeRing = GetCategoryNode(ThisApplication, categNameRing);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория {categNameKonNip2} не найдена. Ошибка!"); }

            else if (listCategoryNodeRing == null) { MessageBox.Show($"Категория {categNameRing} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g23);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, konNip2);

                    ContentFamiliesEnumerator famsRing = listCategoryNodeRing.Families;
                    ContentFamily famRing = GetContentFamily(famsRing, ring);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g23} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {konNip2} не найдено. Ошибка!"); }

                    else if (famRing == null) { MessageBox.Show($"Семейство {ring} не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();
                        MemberManagerErrorsEnum fr;
                        string fm;

                        //string adressNip = "D:/Users/CADExp/Desktop/ДИПЛОМ_ВКР/Парам дет/ГОСТ 23355-78 Исполнение 1.ipt";
                        //string adressKor = "C:/Users/Coddy/Desktop/Глеб/Программная часть/Итог детали/Соединение с коническим ниппелем/ГОСТ 22525-77 Концы корпусных деталей с углом конуса 24°.ipt";
                        //string adressKor = "D:/Users/CADexp/Desktop/Итог детали/Соединение с шаровым ниппелем/ГОСТ 22525-77 Концы корпусных деталей.ipt";

                        //ThisApplication.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, adressKor);
                        //ComponentOccurrence ModelKor = oAssDoc.ComponentDefinition.Occurrences.Add(adressKor, oPositionMatrix);
                        //Inventor.ControlDefinition ctrlDeff2 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff2.Execute();

                        //Добавление из библиотеки компонентов

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow konNip2TableRow = GetTablelRow(rowsNip, "элемент", konNip2RowName);
                        var FileNip = famNip.CreateMember(konNip2TableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip, oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g23RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);
                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        ContentTableRows rowsRing = famRing.TableRows;
                        ContentTableRow ringTableRow = GetTablelRow(rowsRing, "элемент", ringRowName);
                        var FileRing = famRing.CreateMember(ringTableRow, out fr, out fm);
                        ComponentOccurrence ModelRing = oAssDoc.ComponentDefinition.Occurrences.Add(FileRing, oPositionMatrix);
                        ModelRing.Grounded = false;

                        //Собираем все в одно

                        Face oFaceGMOs, oFaceGMKas, oFaceNipOs, oFaceNipVs, oFaceNipKas, oFaceKorOs, oFaceKorVs, oFaceNipTr, oFaceNipTan, oFaceRing, oFaceNipRin;

                        oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipTr = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipTan = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceRing = ModelRing.SurfaceBodies[1].Faces[1];
                        oFaceNipRin = ModelNip.SurfaceBodies[1].Faces[1];



                        string test1, test2;
                        //Ring Nip
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{57772619-3445-C577-DEDD-A5C9E9573E19}")
                            {
                                oFaceNipRin = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ring
                        for (int i = 1; i <= ModelRing.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelRing.SurfaceBodies[1].Faces[i].InternalName == "{577C7664-1B32-4B38-006B-1D6C3E56AB0E}")
                            {
                                oFaceRing = ModelRing.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                       
                        //Гайка ось
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{D6C563AD-11A3-CAA5-4F14-A4749FC2814A}")
                            {
                                oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка кас
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{8EF6084F-251D-2041-F149-6900BC6D3FD8}")
                            {
                                oFaceGMKas = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель ось
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{29347A13-1984-3A98-1C95-835915682BC3}")
                            {
                                oFaceNipOs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель кас {FB243FD5-813F-3C44-D216-E79E6C6BE86D}
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{7CF737AC-B282-926F-2D47-67B0EF1C05A4}")
                            {
                                oFaceNipKas = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Tangency nip {FB243FD5-813F-3C44-D216-E79E6C6BE86D}
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{FB243FD5-813F-3C44-D216-E79E6C6BE86D}")
                            {
                                oFaceNipTan = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель вставка
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{6767762D-5FC3-17F0-47B6-7E540FBB4388}")
                            {
                                oFaceNipVs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель для трубы
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{53A91A09-9AB3-FA79-D00A-91562FBE6BDD}")
                            {
                                oFaceNipTr = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }

                        //Зависимости
                        if (checkBox10.Checked == true)
                        {
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            //mate2 = oAssCompDef.Constraints.AddMateConstraint(faceKon2Os, oFaceGMOs, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);

                            //TangentConstraint tan1;
                            //tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGMKas, oFaceNipKas, true, 0);
                            InsertConstraint insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMKas, oFaceNipOs, true, 0);

                            //Ring
                            oAssCompDef.Constraints.AddMateConstraint(oFaceRing, oFaceNipRin, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            oAssCompDef.Constraints.AddTangentConstraint(oFaceRing, oFaceNipRin, true, 0);

                            //Соединение с готовым
                            MateConstraint mate3 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipTan, faceKon2Vs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceNipTan, faceKon2Vs, true, 0);

                            //InsertConstraint insert1;
                            //insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipVs, faceKon2Vs, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }
                        else if (checkBox9.Checked == true)
                        {
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            //mate2 = oAssCompDef.Constraints.AddMateConstraint(faceKon2Os, oFaceGMOs, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);

                            //TangentConstraint tan1;
                            //tan1 = oAssCompDef.Constraints.AddTangentConstraint(oFaceGMKas, oFaceNipKas, true, 0);
                            InsertConstraint insert2 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMKas, oFaceNipOs, true, 0);
                            //Ring
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(oFaceRing, oFaceNipRin, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            oAssCompDef.Constraints.AddTangentConstraint(oFaceRing, oFaceNipRin, true, 0);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceNipTr, faceTr4, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }
                        oAssemblyDocName = oAssDoc;
                    }
                }
            }
        }



        //По наружному конусу 1 
        private void comboBox6_TextChanged(object sender, EventArgs e)
        {
            string rez = "";
            string dn = "";
            rez = comboBox5.Text;
            dn = comboBox6.Text;
            //comboBox1.Update();
            //comboBox10.Update();
            bool temp1 = false;
            for (int i = 0; i < comboBox6.Items.Count; i++)
            {
                if (Convert.ToString(dn) == Convert.ToString(comboBox6.Items[i]))
                {
                    //comboBox1.Enabled = true;
                    checkBox29.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                //comboBox1.Enabled = false;
                checkBox29.Enabled = false;
                button11.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {

            string rez = "";
            //string dy = "";
            string dn = "";
            //label1.Text = "САС";
            Invoke(new Action(() =>
            {
                //label1.Text = d1;
                //textBox1.Text = d1;
                rez = comboBox7.Text;
                //dy = textBox6.Text;
                dn = comboBox6.Text;
            }));
            rez = comboBox5.Text;
            //dy = textBox7.Text;
            dn = comboBox6.Text;
            int flagNip = 1;
            int flagGM = 1;
            if (dn == "3")
            {
                flagNip = 1;
                flagGM = 1;
            }
            else if (dn == "4")
            {
                flagNip = 2;
                flagGM = 2;
            }
            else if (dn == "6")
            {
                flagNip = 3;
                flagGM = 3;
            }
            else if (dn == "8")
            {
                flagNip = 4;
                flagGM = 4;
            }
            else if (dn == "10")
            {
                flagNip = 5;
                flagGM = 5;
            }
            else if (dn == "12")
            {
                flagNip = 6;
                flagGM = 6;
            }
            else if (dn == "14")
            {
                flagNip = 7;
                flagGM = 7;
            }
            else if (dn == "16")
            {
                flagNip = 8;
                flagGM = 8;
            }
            else if (dn == "18")
            {
                flagNip = 9;
                flagGM = 9;
            }
            else if (dn == "20")
            {
                flagNip = 10;
                flagGM = 10;
            }
            else if (dn == "22")
            {
                flagNip = 11;
                flagGM = 11;
            }
            else if (dn == "25")
            {
                flagNip = 12;
                flagGM = 12;
            }
            else if (dn == "28")
            {
                flagNip = 13;
                flagGM = 13;
            }
            else if (dn == "30")
            {
                flagNip = 14;
                flagGM = 14;
            }
            else if (dn == "32")
            {
                flagNip = 15;
                flagGM = 15;
            }
            else if (dn == "34")
            {
                flagNip = 16;
                flagGM = 16;
            }
            else if (dn == "36")
            {
                flagNip = 17;
                flagGM = 17;
            }
            else if (dn == "38")
            {
                flagNip = 18;
                flagGM = 18;
            }
            

            if (flagNip >= 10)
            {
                nipPK1RowName = $"ГОСТ 13956-74 Исполнение 1-{flagNip}";
            }
            else
            {
                nipPK1RowName = $"ГОСТ 13956-74 Исполнение 1-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g2RowName = $"ГОСТ 13957-74 ПО НАРУЖНОМУ КОНУСУ-{flagGM}";
            }
            else
            {
                g2RowName = $"ГОСТ 13957-74 ПО НАРУЖНОМУ КОНУСУ-0{flagGM}";
            }



            //Ишу нужную категорию
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, categNameGM1);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, categNameNipPK1);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM1} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория {categNameNipPK1} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g2);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, nipPK1);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g2} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {nipPK1} не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();
                        MemberManagerErrorsEnum fr;
                        string fm;

                        //string adressNip = "D:/Users/CADExp/Desktop/ДИПЛОМ_ВКР/Парам дет/ГОСТ 23355-78 Исполнение 1.ipt";
                        //string adressKor = "C:/Users/Глеб/Desktop/Универ/4 курс/ВКР_Диплом/Программная часть/Итог детали/Ниппели для присоединения по наружному конусу/Труба.ipt";
                        //string adressKor = "D:/Users/CADexp/Desktop/Итог детали/Соединение с шаровым ниппелем/ГОСТ 22525-77 Концы корпусных деталей.ipt";

                        //ThisApplication.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, adressKor);
                        //ComponentOccurrence ModelKor = oAssDoc.ComponentDefinition.Occurrences.Add(adressKor, oPositionMatrix);
                        //Inventor.ControlDefinition ctrlDeff2 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff2.Execute();

                        //Добавление из библиотеки компонентов

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow konNip2TableRow = GetTablelRow(rowsNip, "элемент", nipPK1RowName);
                        var FileNip = famNip.CreateMember(konNip2TableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip, oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g2RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);

                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        //Собираем все в одно
                        //Поиск Faces
                        Face oFaceGMOs, oFaceGMVs, oFaceNipOs, oFaceNipVs, oFaceNipKas,  oFaceKorOs, oFaceKorKas;

                        oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorKas = ModelGM.SurfaceBodies[1].Faces[1];

                        //Гайка ось
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{9D4B9E20-9A30-556B-B90D-5E25A8F109D8}")
                            {
                                oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка вставка
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{58F6CA8B-22BE-C97A-3CCA-B159AA91188A}")
                            {
                                oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель ось {3C35ECC3-C044-F4DD-2A98-F52200E91EFF}
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{8E4E222B-6C2A-F58A-82E8-4169B6522888}")
                            {
                                oFaceNipOs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель кас {56926CC1-523D-6BEF-657F-1A34EC5A2E29}
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{56926CC1-523D-6BEF-657F-1A34EC5A2E29}")
                            {
                                oFaceNipKas = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель вставка
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{BCF279FD-C548-76E3-DC65-035196FEF9B8}")
                            {
                                oFaceNipVs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Труба ось
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{6A11E38B-6E80-83B8-5FA1-6CF1F5D02329}")
                        //    {
                        //        oFaceKorOs = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}
                        ////Труба кас
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{476587B9-62CA-4F46-3C67-00CC4B31426B}")
                        //    {
                        //        oFaceKorKas = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}

                        if (checkBox12.Checked == true)
                        {
                            //Штуцер пока не знаю
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(facePNK1Kas, oFaceNipOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMVs, oFaceNipVs, true, 0);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(facePNK1Kas, oFaceNipKas, true, 0);

                            cancelAll();
                            chooseSelection(false);
                            
                        }
                        else if (checkBox11.Checked == true)
                        {
                            //Труба
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(faceTr5, oFaceNipOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMVs, oFaceNipVs, true, 0);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(faceTr5, oFaceNipKas, true, 0);

                            //oAssCompDef.Constraints.AddInsertConstraint(faceTr5, oFaceNipKas, false, 0);

                            cancelAll();
                            chooseSelection(false);
                            
                        }

                        oAssemblyDocName = oAssDoc;


                    }
                }
            }
        }

        //По наружному конусу 2

        private void comboBox8_TextChanged(object sender, EventArgs e)
        {
            string rez = "";
            string dn = "";
            rez = comboBox9.Text;
            dn = comboBox8.Text;
            //comboBox1.Update();
            //comboBox10.Update();
            bool temp1 = false;
            for (int i = 0; i < comboBox8.Items.Count; i++)
            {
                if (Convert.ToString(dn) == Convert.ToString(comboBox8.Items[i]))
                {
                    //comboBox1.Enabled = true;
                    checkBox32.Enabled = true;
                    temp1 = true;
                    break;
                }
            }
            if (temp1 == false)
            {
                //comboBox1.Enabled = false;
                checkBox32.Enabled = false;
                button12.Enabled = false;
            }
        }


        private void button12_Click(object sender, EventArgs e)
        {
            string rez = "";
            //string dy = "";
            string dn = "";
            //label1.Text = "САС";
            Invoke(new Action(() =>
            {
                //label1.Text = d1;
                //textBox1.Text = d1;
                rez = comboBox9.Text;
                //dy = textBox6.Text;
                dn = comboBox8.Text;
            }));
            rez = comboBox9.Text;
            //dy = textBox7.Text;
            dn = comboBox8.Text;
            int flagNip = 1;
            int flagGM = 1;
            if (dn == "3")
            {
                flagNip = 1;
                flagGM = 1;
            }
            else if (dn == "4")
            {
                flagNip = 2;
                flagGM = 2;
            }
            else if (dn == "6")
            {
                flagNip = 3;
                flagGM = 3;
            }
            else if (dn == "8")
            {
                flagNip = 4;
                flagGM = 4;
            }
            else if (dn == "10")
            {
                flagNip = 5;
                flagGM = 5;
            }
            else if (dn == "12")
            {
                flagNip = 6;
                flagGM = 6;
            }
            else if (dn == "14")
            {
                flagNip = 7;
                flagGM = 7;
            }
            else if (dn == "16")
            {
                flagNip = 8;
                flagGM = 8;
            }
            else if (dn == "18")
            {
                flagNip = 9;
                flagGM = 9;
            }
            else if (dn == "20")
            {
                flagNip = 10;
                flagGM = 10;
            }
            else if (dn == "22")
            {
                flagNip = 11;
                flagGM = 11;
            }
            else if (dn == "25")
            {
                flagNip = 12;
                flagGM = 12;
            }
            else if (dn == "28")
            {
                flagNip = 13;
                flagGM = 13;
            }
            else if (dn == "30")
            {
                flagNip = 14;
                flagGM = 14;
            }
            else if (dn == "32")
            {
                flagNip = 15;
                flagGM = 15;
            }
            else if (dn == "34")
            {
                flagNip = 16;
                flagGM = 16;
            }
            else if (dn == "36")
            {
                flagNip = 17;
                flagGM = 17;
            }
            else if (dn == "38")
            {
                flagNip = 18;
                flagGM = 18;
            }


            if (flagNip >= 10)
            {
                nipPK2RowName = $"ГОСТ 13956-74 Исполнение 1-{flagNip}";
            }
            else
            {
                nipPK2RowName = $"ГОСТ 13956-74 Исполнение 1-0{flagNip}";
            }

            if (flagGM >= 10)
            {
                g2RowName = $"ГОСТ 13957-74 ПО НАРУЖНОМУ КОНУСУ-{flagGM}";
            }
            else
            {
                g2RowName = $"ГОСТ 13957-74 ПО НАРУЖНОМУ КОНУСУ-0{flagGM}";
            }

            //Ишу нужную категорию
            ContentTreeViewNode listCategoryNode = GetCategoryNode(ThisApplication, categNameGM1);
            ContentTreeViewNode listCategoryNodeNip = GetCategoryNode(ThisApplication, categNameNipPK2);
            if (listCategoryNode == null)
            {
                MessageBox.Show($"Категория {categNameGM1} не найдена. Ошибка!");
            }
            else if (listCategoryNodeNip == null) { MessageBox.Show($"Категория {categNameNipPK2} не найдена. Ошибка!"); }

            else
            {
                AssemblyDocument oAssDoc = ThisApplication.ActiveDocument as AssemblyDocument;
                if (oAssDoc == null)
                {
                    MessageBox.Show("Сборка не является актинвым документом. Ошибка!");
                }

                //Ищу конкретную деталь в категории
                else
                {
                    ContentFamiliesEnumerator famsGM = listCategoryNode.Families;
                    ContentFamily famGM = GetContentFamily(famsGM, g2);

                    ContentFamiliesEnumerator famsNip = listCategoryNodeNip.Families;
                    ContentFamily famNip = GetContentFamily(famsNip, nipPK2);

                    if (famGM == null)
                    {
                        MessageBox.Show($"Семейство {g2} не найдено. Ошибка!");
                    }
                    else if (famNip == null) { MessageBox.Show($"Семейство {nipPK2} не найдено. Ошибка!"); }


                    else
                    {
                        //Понеслась сборка
                        AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
                        TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
                        Inventor.Matrix oPositionMatrix = aTransGeom.CreateMatrix();
                        MemberManagerErrorsEnum fr;
                        string fm;

                        //string adressNip = "D:/Users/CADExp/Desktop/ДИПЛОМ_ВКР/Парам дет/ГОСТ 23355-78 Исполнение 1.ipt";
                        //string adressKor = "C:/Users/Coddy/Desktop/Глеб/Программная часть/Итог детали/Ниппели для присоединения по наружному конусу/Труба.ipt";
                        //string adressKor = "D:/Users/CADexp/Desktop/Итог детали/Соединение с шаровым ниппелем/ГОСТ 22525-77 Концы корпусных деталей.ipt";

                        //ThisApplication.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, adressKor);
                        //ComponentOccurrence ModelKor = oAssDoc.ComponentDefinition.Occurrences.Add(adressKor, oPositionMatrix);
                        //Inventor.ControlDefinition ctrlDeff2 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff2.Execute();

                        //Добавление из библиотеки компонентов

                        ContentTableRows rowsNip = famNip.TableRows;
                        ContentTableRow konNip2TableRow = GetTablelRow(rowsNip, "элемент", nipPK2RowName);
                        var FileNip = famNip.CreateMember(konNip2TableRow, out fr, out fm);
                        ComponentOccurrence ModelNip = oAssDoc.ComponentDefinition.Occurrences.Add(FileNip, oPositionMatrix);
                        ModelNip.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff0 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff0.Execute();

                        ContentTableRows rows = famGM.TableRows;
                        ContentTableRow gmTableRow = GetTablelRow(rows, "элемент", g2RowName);
                        var FileGM = famGM.CreateMember(gmTableRow, out fr, out fm);

                        ComponentOccurrence ModelGM = oAssDoc.ComponentDefinition.Occurrences.Add(FileGM, oPositionMatrix);
                        ModelGM.Grounded = false;
                        //Inventor.ControlDefinition ctrlDeff1 = ThisApplication.CommandManager.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //ctrlDeff1.Execute();

                        //Собираем все в одно
                        //Поиск Faces
                        Face oFaceGMOs, oFaceGMVs, oFaceNipOs, oFaceNipVs, oFaceNipKas, oFaceKorOs, oFaceKorKas;

                        oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipVs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceNipKas = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorOs = ModelGM.SurfaceBodies[1].Faces[1];
                        oFaceKorKas = ModelGM.SurfaceBodies[1].Faces[1];

                        //Гайка ось
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{9D4B9E20-9A30-556B-B90D-5E25A8F109D8}")
                            {
                                oFaceGMOs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Гайка вставка
                        for (int i = 1; i <= ModelGM.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelGM.SurfaceBodies[1].Faces[i].InternalName == "{58F6CA8B-22BE-C97A-3CCA-B159AA91188A}")
                            {
                                oFaceGMVs = ModelGM.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель ось
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{3C35ECC3-C044-F4DD-2A98-F52200E91EFF}")
                            {
                                oFaceNipOs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель кас
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{56926CC1-523D-6BEF-657F-1A34EC5A2E29}")
                            {
                                oFaceNipKas = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Ниппель вставка
                        for (int i = 1; i <= ModelNip.SurfaceBodies[1].Faces.Count; i++)
                        {
                            if (ModelNip.SurfaceBodies[1].Faces[i].InternalName == "{BCF279FD-C548-76E3-DC65-035196FEF9B8}")
                            {
                                oFaceNipVs = ModelNip.SurfaceBodies[1].Faces[i] as Face;
                                break;
                            }
                        }
                        //Труба ось
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{6A11E38B-6E80-83B8-5FA1-6CF1F5D02329}")
                        //    {
                        //        oFaceKorOs = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}
                        //Труба кас
                        //for (int i = 1; i <= ModelKor.SurfaceBodies[1].Faces.Count; i++)
                        //{
                        //    if (ModelKor.SurfaceBodies[1].Faces[i].InternalName == "{476587B9-62CA-4F46-3C67-00CC4B31426B}")
                        //    {
                        //        oFaceKorKas = ModelKor.SurfaceBodies[1].Faces[i] as Face;
                        //        break;
                        //    }
                        //}

                        if (checkBox14.Checked == true)
                        {
                            //Штуцер пока не готово
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(facePNK2Kas, oFaceNipOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMVs, oFaceNipVs, true, 0);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(facePNK2Kas, oFaceNipKas, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }
                        else if (checkBox13.Checked == true)
                        {
                            //Труба
                            //Зависимости
                            MateConstraint mate1, mate2;
                            mate1 = oAssCompDef.Constraints.AddMateConstraint(oFaceNipOs, oFaceGMOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            mate2 = oAssCompDef.Constraints.AddMateConstraint(faceTr6, oFaceNipOs, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);

                            InsertConstraint insert1;
                            insert1 = oAssCompDef.Constraints.AddInsertConstraint(oFaceGMVs, oFaceNipVs, true, 0);

                            TangentConstraint tan1;
                            tan1 = oAssCompDef.Constraints.AddTangentConstraint(faceTr6, oFaceNipKas, true, 0);

                            cancelAll();
                            chooseSelection(false);
                        }
                        

                        oAssemblyDocName = oAssDoc;


                    }
                }
            }
        }

        //Кнопки Отменить

        public void cancelAll()
        {
            //sh1
            comboBox1.Text = "";
            comboBox10.Text = "";
            comboBox15.Text = "";
            comboBox1.Enabled = false;
            comboBox10.Enabled = false;
            comboBox15.Enabled = false;
            button5.Enabled = false;
            checkBox2.Enabled = false;
            checkBox1.Enabled = false;
            checkBox15.Enabled = false;
            checkBox2.Checked = false;
            checkBox1.Checked = false;
            checkBox15.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;

            flagG = 0;
            flagGR = 0;
            faceTr1bl = false;
            flagSh1faceVs = false;

            //sh2
            comboBox11.Text = "";
            comboBox4.Text = "";
            comboBox16.Text = "";
            comboBox11.Enabled = false;
            comboBox4.Enabled = false;
            comboBox16.Enabled = false;
            button8.Enabled = false;
            checkBox16.Enabled = false;
            checkBox17.Enabled = false;
            checkBox18.Enabled = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;

            flagP = 0;
            flagPR = 0;
            faceTr2bl = false;
            flagSh2faceVs = false;

            //kn1
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox12.Text = "";
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox12.Enabled = false;

            button7.Enabled = false;

            checkBox19.Enabled = false;
            checkBox20.Enabled = false;
            checkBox21.Enabled = false;
            checkBox22.Enabled = false;

            checkBox19.Checked = false;
            checkBox20.Checked = false;
            checkBox21.Checked = false;
            checkBox22.Checked = false;

            checkBox8.Checked = false;
            checkBox7.Checked = false;

            flagK1 = 0;
            flagKDn1 = 0;
            flagK1R = 0;
            faceTr3bl = false;
            flagKon1faceVs = false;

            //kn2
            comboBox13.Text = "";
            comboBox14.Text = "";
            comboBox5.Text = "";
            comboBox13.Enabled = false;
            comboBox14.Enabled = false;
            comboBox5.Enabled = false;

            button10.Enabled = false;

            checkBox23.Enabled = false;
            checkBox24.Enabled = false;
            checkBox25.Enabled = false;
            checkBox26.Enabled = false;

            checkBox23.Checked = false;
            checkBox24.Checked = false;
            checkBox25.Checked = false;
            checkBox26.Checked = false;

            checkBox9.Checked = false;
            checkBox10.Checked = false;

            flagK2 = 0;
            flagKDn = 0;
            flagK2R = 0;
            flagKon2faceVs = false;
            faceTr4bl = false;

            //pnk1
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox6.Enabled = false;
            comboBox7.Enabled = false;
            button11.Enabled = false;
            checkBox29.Enabled = false;
            checkBox27.Enabled = false;
            checkBox28.Enabled = false;
            checkBox29.Checked = false;
            checkBox27.Checked = false;
            checkBox28.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;

            flagPNK1 = 0;
            flagPNK1R = 0;
            flagPNK1faceKas = false;
            flagPNK1faceOs = false;
            faceTr5bl = false;

            //pnk2
            comboBox8.Text = "";
            comboBox9.Text = "";
            comboBox8.Enabled = false;
            comboBox9.Enabled = false;
            button12.Enabled = false;
            checkBox30.Enabled = false;
            checkBox31.Enabled = false;
            checkBox32.Enabled = false;
            checkBox30.Checked = false;
            checkBox31.Checked = false;
            checkBox32.Checked = false;
            checkBox14.Checked = false;
            checkBox13.Checked = false;

            flagPNK2 = 0;
            flagPNK2R = 0;
            flagPNK2faceKas = false;
            flagPNK2faceOs = false;
            faceTr6bl = false;

            chooseSelection(false);

        }

        private void button37_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            comboBox10.Text = "";
            comboBox16.Text = "";
            comboBox1.Enabled = false;
            comboBox10.Enabled = false;
            comboBox16.Enabled = false;
            button5.Enabled = false;
            checkBox2.Enabled = false;
            checkBox1.Enabled = false;
            checkBox15.Enabled = false;
            checkBox2.Checked = false;
            checkBox1.Checked = false;
            checkBox15.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;

            flagG = 0;
            flagGR = 0;
            faceTr1bl = false;
            flagSh1faceVs = false;

            chooseSelection(false);
        }

        private void button41_Click(object sender, EventArgs e)
        {
            comboBox11.Text = "";
            comboBox4.Text = "";
            comboBox15.Text = "";
            comboBox11.Enabled = false;
            comboBox4.Enabled = false;
            comboBox15.Enabled = false;
            button8.Enabled = false;
            checkBox16.Enabled = false;
            checkBox17.Enabled = false;
            checkBox18.Enabled = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;

            flagP = 0;
            flagPR = 0;
            faceTr2bl = false;
            flagSh2faceVs = false;

            chooseSelection(false);
        }

        private void button42_Click(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox12.Text = "";
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox12.Enabled = false;

            button7.Enabled = false;

            checkBox19.Enabled = false;
            checkBox20.Enabled = false;
            checkBox21.Enabled = false;
            checkBox22.Enabled = false;

            checkBox19.Checked = false;
            checkBox20.Checked = false;
            checkBox21.Checked = false;
            checkBox22.Checked = false;

            checkBox8.Checked = false;
            checkBox7.Checked = false;

            flagK1 = 0;
            flagKDn1 = 0;
            flagK1R = 0;
            faceTr3bl = false;
            flagKon1faceVs = false;

            chooseSelection(false);
        }

        private void button45_Click(object sender, EventArgs e)
        {
            comboBox13.Text = "";
            comboBox14.Text = "";
            comboBox5.Text = "";
            comboBox13.Enabled = false;
            comboBox14.Enabled = false;
            comboBox5.Enabled = false;

            button10.Enabled = false;

            checkBox23.Enabled = false;
            checkBox24.Enabled = false;
            checkBox25.Enabled = false;
            checkBox26.Enabled = false;

            checkBox23.Checked = false;
            checkBox24.Checked = false;
            checkBox25.Checked = false;
            checkBox26.Checked = false;

            checkBox9.Checked = false;
            checkBox10.Checked = false;

            flagK2 = 0;
            flagKDn = 0;
            flagK2R = 0;
            flagKon2faceVs = false;
            faceTr4bl = false;

            chooseSelection(false);
        }

        private void button46_Click(object sender, EventArgs e)
        {
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox6.Enabled = false;
            comboBox7.Enabled = false;
            button11.Enabled = false;
            checkBox29.Enabled = false;
            checkBox27.Enabled = false;
            checkBox28.Enabled = false;
            checkBox29.Checked = false;
            checkBox27.Checked = false;
            checkBox28.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;

            flagPNK1 = 0;
            flagPNK1R = 0;
            flagPNK1faceKas = false;
            flagPNK1faceOs = false;
            faceTr5bl = false;

            chooseSelection(false);
        }

        private void button44_Click(object sender, EventArgs e)
        {
            comboBox8.Text = "";
            comboBox9.Text = "";
            comboBox8.Enabled = false;
            comboBox9.Enabled = false;
            button12.Enabled = false;
            checkBox30.Enabled = false;
            checkBox31.Enabled = false;
            checkBox32.Enabled = false;
            checkBox30.Checked = false;
            checkBox31.Checked = false;
            checkBox32.Checked = false;
            checkBox14.Checked = false;
            checkBox13.Checked = false;

            flagPNK2 = 0;
            flagPNK2R = 0;
            flagPNK2faceKas = false;
            flagPNK2faceOs = false;
            faceTr6bl = false;

            chooseSelection(false);
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            if (panel3.Visible == false)
            {
                panel3.Visible = true;
                panel4.Visible = true;
            }
            else
            {
                panel3.Visible = false;
                panel4.Visible = false;
            }                
        }

        private void label30_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage1;
            label1.Text = "Ниппельное соединение. ГОСТ 23355-78 Шаровой ниппель. Исполнение 1";
            cancelAll();
        }

        private void label31_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage2;
            label1.Text = "Ниппельное соединение. ГОСТ 23355-78 Шаровой ниппель. Исполнение 2";
            cancelAll();
        }

        private void label32_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage3;
            label1.Text = "Ниппельное соединение. ГОСТ 28016-89 Конический ниппель. Исполнение 1";
            cancelAll();
        }

        private void label33_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage4;
            label1.Text = "Ниппельное соединение. ГОСТ 28016-89 Конический ниппель. Исполнение 2";
            cancelAll();
        }

        private void label34_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage5;
            label1.Text = "Ниппельное соединение. Ниппель ГОСТ 17956-74. Исполнение 1";
            cancelAll();
        }

        private void label35_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            tabControl1.SelectedTab = tabPage6;
            label1.Text = "Ниппельное соединение. Ниппель ГОСТ 17956-74. Исполнение 2";
            cancelAll();
        }
    }
}


