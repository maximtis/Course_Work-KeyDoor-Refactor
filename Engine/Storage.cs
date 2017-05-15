using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Threading.Tasks;
using _3Course_Program_Lab_5;
using Word = Microsoft.Office.Interop.Word;

/*
Варіант 10.
    Кодовий замок.Потрібно розробити засобами Rational Rose модель програмного забезпечення кодового замка,
що регулює доступ в приміщення.Кодовий замок складається з панелі з кнопками (цифри «0» ... «9», кнопка «Виклик», кнопка «Контроль»),
цифрового дисплея, електромеханічного замку, дзвінка.
    Панель з кнопками встановлюється із зовнішнього боку дверей,
замок встановлюється з внутрішньої сторони дверей, дзвінок встановлюється всередині приміщення, що охороняється.
У звичайному стані замок закритий.Доступ до приміщення здійснюється після набору коду доступу, що складається з чотирьох цифр.
    Під час набору коду введені цифри відображаються на дисплеї.
Якщо код набрано правильно, то замок відкривається на деякий час, 
після чого двері знову закриваються.Вміст дисплея очищається.
    Кнопка «Виклик» використовується
для подачі звукового сигналу всередині приміщення.
    Кнопка «Контроль» використовується для зміни кодів.
Зміна коду доступу здійснюється наступним чином.
    При відкритій двері потрібно набрати код контролю, що складається з чотирьох цифр,
і новий код доступу. 
Для зміни коду контролю потрібно при  
відкритій двері і утримуючи кнопку «Виклик» набрати код контролю, після чого - новий код контролю.
*/
namespace _3Course_Program_Lab_5.Engine
{
    static class CommandDictionary
    {
        public static Dictionary<string, int> CDictionary = new Dictionary<string, int>();
        static CommandDictionary()
        {
            CDictionary.Add("Control", 1);
            CDictionary.Add("Call", 2);
        }
    }
    class Storage
    {
        public Storage()
        {
            MasterPassword = "0000";
            CurrentPassword = "1111";
        }
        public string CurrentPassword { get; set; }
        public string MasterPassword { get; set; }
    }

    abstract class ServiceBase
    {
        public abstract void RunService();
        public abstract void StopService();
        public abstract void Reset();
    }

    class MonitoringService: ServiceBase
    {
        public string Log { get; set; }
        public MonitoringService()
        {
            ServiceTimer = new DispatcherTimer();
            ServiceTimer.Interval = TimeSpan.FromSeconds(1);
            ServiceTimer.Tick += ServiceTimer_Tick;
        }
        private void ServiceTimer_Tick(object sender, EventArgs e)
        {
            ControlTime++;
        }
        public void HandleReset(object sender, EventArgs e)
        {
            Reset();
        }
        public int controlTime;
        public int ControlTime
        {
            get
            {
                return controlTime;
            }
            set
            {
                controlTime = value;
                if (controlTime == 5)
                    if (TimeOut != null)
                        TimeOut(this, new EventArgs());
            }
        }
        public event EventHandler TimeOut;
        public DispatcherTimer ServiceTimer { get; set; }
        public override void RunService()
        {
            ServiceTimer.Start();
        }
        public override void Reset()
        {
            ControlTime = 0;
        }
        public override void StopService()
        {
            ServiceTimer.Stop();
        }
    }

    class Journal
    {
        private Word.Documents wordDocs;
        private Word.Document wordDoc;
        private Word.Application wordapp = new Word.Application();

        public Journal()
        {
            wordapp.Visible = false;
            Object template = Type.Missing;
            Object newTemplate = false;
            Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
            wordDocs = wordapp.Documents;
            wordDoc = wordDocs[1];

            Object fileName = @"C:\Users\zao.CCD\Desktop\3Course-Program-Lab-5\Log.doc";
            Object fileFormat = Word.WdSaveFormat.wdFormatDocument;
            Object lockComments = false;
            Object password = "";
            Object addToRecent = false;
            Object writePassword = "";
            Object readOnly = false;
            Object embedTrueFonts = false;
            Object saveNativePictureFormat = false;
            Object saveFormsData = false;
            Object saveAsAOCLetter = Type.Missing;
            Object encoding = Type.Missing;
            Object insertLineBreaks = Type.Missing;
            Object allowSubstitutions = Type.Missing;
            Object lineEnding = Type.Missing;
            Object addBiDiMarks = Type.Missing;
            Object compatibility = Type.Missing;

            wordDoc.SaveAs2(ref fileName, ref fileFormat, ref lockComments,
                            ref password, ref addToRecent, ref writePassword,
                            ref readOnly, ref embedTrueFonts, ref saveNativePictureFormat,
                            ref saveFormsData, ref saveAsAOCLetter, ref encoding,
                            ref insertLineBreaks, ref allowSubstitutions, ref lineEnding,
                            ref addBiDiMarks, ref compatibility);
            wordDoc.Close();
            
        }
        public void WriteToWord(string text)
        {
            Word.Application ap = new Word.Application();
            try
            {

                Word.Document doc = ap.Documents.Open(@"C:\Users\zao.CCD\Desktop\3Course-Program-Lab-5\Log.doc", ReadOnly: false, Visible: false);
                doc.Activate();

                Word.Selection sel = ap.Selection;

                if (sel != null)
                {
                    switch (sel.Type)
                    {
                        case Word.WdSelectionType.wdSelectionIP:
                            sel.TypeText(DateTime.Now.ToString());
                            sel.TypeParagraph();
                            sel.TypeText(text);
                            sel.TypeParagraph();
                            break;

                        default:
                            Console.WriteLine("Selection type not handled; no writing done");
                            break;

                    }

                    // Remove all meta data.
                    doc.RemoveDocumentInformation(Word.WdRemoveDocInfoType.wdRDIAll);

                    ap.Documents.Save(NoPrompt: true, OriginalFormat: true);
                }
                else
                {
                    Console.WriteLine("Unable to acquire Selection...no writing to document done..");
                }

                ap.Documents.Close(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Caught: " + ex.Message); // Could be that the document is already open (/) or Word is in Memory(?)
            }
            finally
            {
                ((Word._Application)ap).Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ap);
            }
        }

    }
    class PasswordControl
    {
        public PasswordControl()
        {
            PasswordStorage = new Storage();
        }
        public Storage PasswordStorage { get; private set; }
        public void ChangePassword(string NewPassword)
        {
                PasswordStorage.CurrentPassword = NewPassword;
        }
        public void ChangeMasterPassword(string NewPassword)
        {
                PasswordStorage.MasterPassword = NewPassword;
        }
    }

    enum WorkStates { Standby=0, Open=1, Control=2, Admin=3 ,Changing=4 };
    enum PassType { Door = 1, Master = 2 };
    class CodeLockEngine
    {
        TextBox TPanel;
        PassType passType;
        public WorkStates State { get; private set; }
        public void OpenDoor()
        {
            State = WorkStates.Open;
            PasswordBuffer = string.Empty;
            TPanel.Text = "Open";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Door was Open";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Switch Lock to state 'Open' ";
        }
        public void ControlMode()
        {
            State = WorkStates.Control;
            PasswordBuffer = string.Empty;
            TPanel.Text = "Control";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Entered corrct admin pass";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Switch Lock to state 'Control' ";
        }
        public void AdminMode()
        {
            State = WorkStates.Admin;
            MasterPasswordBuffer = string.Empty;
            TPanel.Text = "Admin";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Pressed admin button";
            Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Switch Lock to state 'Admin' ";
        }
        public void Changing(PassType type)
        {
            passType = type;
            State = WorkStates.Changing;
            PasswordBuffer = string.Empty;
            switch (type)
            {
                case PassType.Door:
                    TPanel.Text = "New Password";
                    Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Password was changed";
                    Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "New password for door was entered ";
                    break;
                case PassType.Master:
                    Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "Password was changed";
                    Monitor.Log += Environment.NewLine + " " + DateTime.Now + " : " + "New admin password was entered ";
                    TPanel.Text = "New Master Password";
                    break;
            }
        }
        public void End()
        {
            State = WorkStates.Standby;
            PasswordBuffer = string.Empty;
            TPanel.Text = string.Empty;
        }
        public void Error()
        {
            PasswordBuffer = string.Empty;
            Message("Error");
            State = WorkStates.Standby;
        }
        public void Message(string m)
        {
            TPanel.Text = m;
        }

        string passwordBuffer;
        public string PasswordBuffer
        {
            get
            {
                return passwordBuffer;
            }
            set
            {
                TPanel.Text = value;
                passwordBuffer = value;
            }
        }
        public string MasterPasswordBuffer { get; private set; }
        public CodeLockEngine( TextBox t)
        {
            Monitor = new MonitoringService();
            Monitor.TimeOut += Monitor_TimeOut;
            Avalible += Monitor.HandleReset;
            Monitor.RunService();
            PasswordControl = new PasswordControl();
            TPanel = t;
        }

        private void Monitor_TimeOut(object sender, EventArgs e)
        {
            End();
        }
        public MonitoringService Monitor{ get; set; }
        public PasswordControl PasswordControl { get; set; }

        public event EventHandler Avalible;
        private bool CheckingPassword()
        {
            return PasswordBuffer == PasswordControl.PasswordStorage.CurrentPassword;
        }
        private int CheckingNumberButton(object number)
        {
            dynamic value = null;
            if (int.TryParse(number.ToString(), value))
                return value;
            throw new InvalidCastException("Invalid Command");
        }
        private bool IsNumberButton(object sender)
        {
            int ConrolValue;
            return int.TryParse((sender as Button).Content.ToString(), out ConrolValue);
        }
        private bool IsControlButton(object sender)
        {
            return ((string)(sender as Button).Content) == "Control";
        }
        private bool IsCallButton(object sender)
        {
            return ((string)(sender as Button).Content) == "Call";
        }

        private void ChangeState(WorkStates NewState)
        {
            State = NewState;
        }
        public void ButtonProcessor(object sender, RoutedEventArgs e)
        {
            if (Avalible != null)
                Avalible(this, new EventArgs());
            try {
                switch (State)
                {
                    case WorkStates.Standby:
                        if(IsNumberButton(sender))
                            PasswordBuffer += (sender as Button).Content;
                        if (PasswordBuffer.Length == 4)
                            if ( PasswordBuffer == PasswordControl.PasswordStorage.CurrentPassword)
                            OpenDoor();
                            else
                                Error();
                        break;

                    case WorkStates.Open:
                        if (IsControlButton(sender))
                            ControlMode();
                        if (IsCallButton(sender))
                            AdminMode();
                        break;

                    case WorkStates.Control:
                        if (IsNumberButton(sender))
                            PasswordBuffer += (sender as Button).Content;
                        if (PasswordBuffer.Length == 4)
                            if (PasswordBuffer == PasswordControl.PasswordStorage.MasterPassword)
                                Changing(PassType.Door);
                            else
                                Error();
                        break;

                    case WorkStates.Admin:
                        if (IsNumberButton(sender))
                            PasswordBuffer += (sender as Button).Content;
                        if (PasswordBuffer.Length == 4)
                            if (PasswordBuffer == PasswordControl.PasswordStorage.MasterPassword)
                                Changing(PassType.Master);
                            else
                                Error();
                        break;

                    case WorkStates.Changing:
                        if (IsNumberButton(sender))
                            PasswordBuffer += (sender as Button).Content;
                        if (PasswordBuffer.Length == 4)
                            switch(passType)
                            {
                                case PassType.Door:
                                    PasswordControl.ChangePassword(PasswordBuffer);
                                    Message("Success");
                                    break;
                                case PassType.Master:
                                    PasswordControl.ChangeMasterPassword(PasswordBuffer);
                                    Message("Success");
                                    break;
                            }
                        break;

                }
            }
            catch(InvalidCastException ex)
            {
                ErrorMessage(ex.Message);
            }
                     
        }
        private void ErrorMessage(string message)
        {
            throw new NotImplementedException();
        }
    }

}
