using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using AppCore = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using AppSystemVariableChangedEventArgs = Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Runtime;

namespace Function
{
    internal static class StartEvents
    {
        private static bool _initialized, _needUpdRibbonDetected;
        private static string ribbonTab = "nCommand", ribbonTabId = "nCommand_Id";        
        //(string, string, string) команда / название / описание
        private static List<(string, List<(string, string, string)>)> commands = new List<(string, List<(string, string, string)>)>()
        {
            ("Вывод", new List<(string, string, string)> 
            { 
                ("bad_tab", "Экспорт в Excel", "Экспорт в Excel таблиц, состоящих из примитивов (линии/полилинии/текст/мультитекст).")
            }),                     
        }; 
        /// <summary>
        /// Инициализация
        /// </summary>
        public static void Initialize()
        {
            if (!_initialized)
            {
                _initialized = true;

                AppCore.Idle += Application_Idle_RibbonUpdate;
                AppCore.SystemVariableChanged += App_SysVarChanged_RibbonUpdate;
            }
        }
        private static void App_SysVarChanged_RibbonUpdate (object sender, AppSystemVariableChangedEventArgs e)
        {
            if (!_needUpdRibbonDetected
                && e.Name.Equals("WSCURRENT",
                StringComparison.OrdinalIgnoreCase)
                || e.Name.Equals("RIBBONSTATE",
                StringComparison.OrdinalIgnoreCase))
            {
                _needUpdRibbonDetected = true;
                AppCore.Idle += Application_Idle_RibbonUpdate;
            }
        }
        private static void Application_Idle_RibbonUpdate(object sender, EventArgs e)
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon != null)
            {
                AppCore.Idle -= Application_Idle_RibbonUpdate;
                _needUpdRibbonDetected = false;
                CreateRibbonTab();
            }
        }
        private static void CreateRibbonTab()
        {
            try
            {
                // Получаем доступ к ленте
                RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
                RibbonTab ribTab = null;
                // добавляем свою вкладку
                foreach (RibbonTab tab in ribCntrl.Tabs)
                {
                    if (tab.Id.Equals(ribbonTabId) & tab.Title.Equals(ribbonTab))
                    {
                        ribTab = tab;
                        break;
                    }
                }
                if (ribTab == null)
                {
                    ribTab = new RibbonTab();
                    ribTab.Title = ribbonTab; // Заголовок вкладки
                    ribTab.Id = ribbonTabId; // Идентификатор вкладки
                    ribCntrl.Tabs.Add(ribTab); // Добавляем вкладку в ленту
                }
                // добавляем содержимое в свою вкладку (одну панель)
                addExampleContent(ribTab);
                // Делаем вкладку активной (не желательно, ибо неудобно)
                //ribTab.IsActive = true;
                // Обновляем ленту (если делаете вкладку активной, то необязательно)
                ribCntrl.UpdateLayout();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }
        // Строим новую панель в нашей вкладке
        private static void addExampleContent(RibbonTab ribTab)
        {
            try
            {
                foreach ((string, List<(string, string, string)>) command in commands)
                {
                    RibbonPanel ribPanel = null;
                    RibbonPanelSource ribSourcePanel = null;
                    // создаем panel source                
                    foreach (RibbonPanel panel in ribTab.Panels)
                    {
                        if (panel.Source.Title.Equals(command.Item1))
                        {
                            ribPanel = panel;
                            ribSourcePanel = panel.Source;
                            break;
                        }
                    }
                    if (ribPanel == null)
                    {
                        // создаем panel source
                        ribSourcePanel = new RibbonPanelSource();
                        ribSourcePanel.Title = command.Item1;
                        ribPanel = new RibbonPanel();
                        ribPanel.Source = ribSourcePanel;
                        ribTab.Panels.Add(ribPanel);
                    }
                    bool stop = true;
                    if (command.Item2.Count > 1)
                    {
                        RibbonSplitButton sb = CreateSplitButton(command.Item2);
                        foreach (Object obj in ribSourcePanel.Items)
                        {
                            if (obj is RibbonSplitButton)
                            {
                                RibbonSplitButton rbc = obj as RibbonSplitButton;
                                if (rbc != null)
                                {
                                    if (rbc.Text.Equals(sb.Text))
                                    {
                                        stop = false;
                                        break;
                                    }
                                }
                            }
                        }
                        if (stop) ribSourcePanel.Items.Add(sb);
                    }
                    else if (command.Item2.Count == 1)
                    {
                        RibbonButton rb = CreateButton(command.Item2[0]);
                        foreach (Object obj in ribSourcePanel.Items)
                        {
                            if (obj is RibbonButton & !(obj is RibbonSplitButton))
                            {
                                RibbonButton rbc = obj as RibbonButton;
                                if (rbc != null)
                                {
                                    if (rbc.CommandParameter.Equals(rb.CommandParameter))
                                    {
                                        stop = false;
                                        break;
                                    }
                                }
                            }
                        }
                        if (stop) ribSourcePanel.Items.Add(rb);
                    }
                    if (stop)
                    {
                        if (ribSourcePanel.Items.Count == 5) ribSourcePanel.Items.Add(new RibbonPanelBreak());
                        else ribSourcePanel.Items.Add(new RibbonRowBreak());
                    }
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }
        /// <summary>
        /// создает кнопку из текста
        /// </summary>
        /// <param name="bp">список из комманды, названия и описания</param>
        /// <returns></returns>
        private static RibbonButton CreateButton((string, string, string) bp)
        {
            RibbonToolTip tt = new RibbonToolTip();
            tt.IsHelpEnabled = false;
            RibbonButton ribBtn = new RibbonButton();
            ribBtn.CommandParameter = tt.Command = bp.Item1;
            ribBtn.Text = tt.Title = ribBtn.Name = bp.Item2;
            ribBtn.CommandHandler = new RibbonCommandHandler();
            ribBtn.Orientation = System.Windows.Controls.Orientation.Horizontal;
            ribBtn.Size = RibbonItemSize.Standard;
            //ribBtn.Image = LoadImage("icon_16");
            ribBtn.ShowImage = false;
            ribBtn.ShowText = true;
            tt.Content = bp.Item3;
            ribBtn.ToolTip = tt;
            return ribBtn;
        }
        private static RibbonSplitButton CreateSplitButton(List<(string, string, string)> bpslist)
        {
            // создаем split button
            RibbonSplitButton risSplitBtn = new RibbonSplitButton();
            /* Для RibbonSplitButton ОБЯЗАТЕЛЬНО надо указать
             * свойство Text, а иначе при поиске команд в автокаде
             * будет вылетать ошибка.
             */
            risSplitBtn.Text = bpslist[0].Item2;
            // Ориентация кнопки
            risSplitBtn.Orientation = System.Windows.Controls.Orientation.Vertical;
            // Размер кнопки
            risSplitBtn.Size = RibbonItemSize.Standard;
            // Показывать изображение
            risSplitBtn.ShowImage = false;
            // Показывать текст
            risSplitBtn.ShowText = true;
            // Стиль кнопки
            risSplitBtn.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            risSplitBtn.ResizeStyle = RibbonItemResizeStyles.NoResize;
            risSplitBtn.ListStyle = RibbonSplitButtonListStyle.List;
            foreach ((string, string, string) bp in bpslist)
            {
                risSplitBtn.Items.Add(CreateButton(bp));
            }
            risSplitBtn.Current = risSplitBtn.Items.First();
            return risSplitBtn;
        }
        // Получение картинки из ресурсов
        // Данная функция найдена на просторах интернет
        //System.Windows.Media.Imaging.BitmapImage LoadImage(string ImageName)
        //{
        //    return new System.Windows.Media.Imaging.BitmapImage(
        //        new Uri("pack://application:,,,/ACadRibbon;component/" + ImageName + ".png"));
        //}
        /* Собственный обраотчик команд
         * Это один из вариантов вызова команды по нажатию кнопки
         */
        private class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }
            public event EventHandler CanExecuteChanged;
            public void Execute(object parameter)
            {
                if (parameter is RibbonButton)
                {
                    // Просто берем команду, записанную в CommandParameter кнопки
                    // и выпоняем её используя функцию SendStringToExecute
                    RibbonButton button = parameter as RibbonButton;
                    acadApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        button.CommandParameter + " ", true, false, true);
                }
            }

        }
    }
    public class ExampleRibbon : IExtensionApplication
    {
        public void Initialize()
        {
            StartEvents.Initialize();
        }        
        public void Terminate()
        {           
        }
    }
}
