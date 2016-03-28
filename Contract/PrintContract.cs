using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows;
using GalaSoft.MvvmLight;

namespace Contract
{
    public class PrintContract : ObservableObject
    {        
        Word._Document document;

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;

        public void WayWord()
        {
            Word._Application application = new Word.Application();
            // создаем путь к файлу 
            Object templatePathObj = @"C:\Users\Evgeniy\Desktop\Договор услуги (диагностика).docx";

            try
            {
                document = application.Documents.Add(ref  templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref  missingObj, ref missingObj);
                application.Quit(ref missingObj, ref  missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;
        }

        public void InsertText(string DateContract, string Name, string NameBaby, string Birthday, string Addres, string Phone, string EMail, string Diagnosis, string Time, string Serviсe, string GetResult, string Many)
        {
            string dateContract = "@@Date", 
                name = "@@Name", 
                baby = "@@Baby", 
                birthday = "@@Birthday", 
                address = "@@Address", 
                phone = "@@Phone", 
                email = "@@Email", 
                diagnosis = "@@Diagnosis", 
                time = "@@@Time", 
                serviсe = "@@Service", 
                getResult = "@@GetResult", 
                many = "@@Many";
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            // обходим все разделы документа
            try
            {
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    // берем всю секцию диапазоном
                    wordRange = document.Sections[i].Range;

                    Word.Find wordFindObj = wordRange.Find;
                    object[] wordName = new object[15] { name, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Name, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordBabyName = new object[15] { baby, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, NameBaby, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordBirthday = new object[15] { birthday, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Birthday, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordAddres = new object[15] { address, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Addres, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordPhone = new object[15] { phone, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Phone, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordDiagnosis = new object[15] { diagnosis, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Diagnosis, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordDate = new object[15] { dateContract, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, DateContract, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordEmail = new object[15] { email, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, EMail, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordTime = new object[15] { time, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Time, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordService = new object[15] { serviсe, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Serviсe, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordGetResult = new object[15] { getResult, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, GetResult, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    object[] wordMany = new object[15] { many, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, Many, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };

                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordName);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordBabyName);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordBirthday);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordAddres);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordPhone);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordDiagnosis);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordDate);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordEmail);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordTime);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordService);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordGetResult);
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordMany);
                }
            }
            catch
            {
                MessageBox.Show("попробуйте еще раз"); //какой-то Error!!! но все равно выполняется)) это магия)
            }            
        }      

    }
}
