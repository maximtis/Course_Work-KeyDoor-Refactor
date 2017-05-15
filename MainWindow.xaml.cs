using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using _3Course_Program_Lab_5.Engine;

namespace _3Course_Program_Lab_5
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        CodeLockEngine codeLock;
        public MainWindow()
        {
            Closing += MainWindow_Closing;
            InitializeComponent();
            codeLock = new CodeLockEngine(InfoPanel);
            button0.Click += codeLock.ButtonProcessor;
            button1.Click += codeLock.ButtonProcessor;
            button2.Click += codeLock.ButtonProcessor;
            button3.Click += codeLock.ButtonProcessor;
            button4.Click += codeLock.ButtonProcessor;
            button5.Click += codeLock.ButtonProcessor;
            button6.Click += codeLock.ButtonProcessor;
            button7.Click += codeLock.ButtonProcessor;
            button8.Click += codeLock.ButtonProcessor;
            button9.Click += codeLock.ButtonProcessor;
            buttonControl.Click += codeLock.ButtonProcessor;
            buttonCall.Click += codeLock.ButtonProcessor;
        }

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Journal journal = new Journal();
            journal.WriteToWord(codeLock.Monitor.Log);
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}
