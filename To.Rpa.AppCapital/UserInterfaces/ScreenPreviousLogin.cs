using DocumentFormat.OpenXml.Packaging;
using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace To.Rpa.AppCapital.Interfaces
{
    public class ScreenPreviousLogin : IUserInterface
    {
        public void CheckCompleteLoad()
        {
        }

        public void DoActivities()
        {
            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void DoOperations()
        {
            Process[] p = Process.GetProcessesByName("pcsws");
            //p[0].WaitForInputIdle();

            Thread.Sleep(3000);

            AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));
            //WinFormAdapter.SetText(_0_Descendants_1[1], "BFPJUARUI");

            ValuePattern etb = _0_Descendants_1[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb.SetValue("BFPROBOP2");

            AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
            //WinFormAdapter.SetText(_0_Descendants_2[1], "BFPJUARUI2");

            ValuePattern etb2 = _0_Descendants_2[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb2.SetValue("BFPROBOP6");

            //WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(_0, "Aceptar"));
            AutomationElementCollection _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Aceptar"));

            var invokePattern = _0_Descendants_3[0].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            /************************** INTENTOS ******************/

            //intento 1
            //var tt = _0_Descendants_3[0].GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
            //invokePattern.Invoke();

            //intento 2
            //AutomationElementCollection xx = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Sesión A - [24 x 80]"));

            //AutomationElement w = xx[0].FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));
            //AutomationElement w1 = xx[0].FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));


            //AutomationProperty[] t = xx[0].GetSupportedProperties();

            /************************** INTENTOS ******************/
            
        }

        //private void modoantiguo() {
        //    //string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx;" + @"Extended Properties='Excel 8.0;HDR=Yes;'";
        //    //string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\LOG\DATA_PROCESO.xlsx;Extended Properties='Excel 8.0;HDR=Yes;'";
        //    string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LOG\DATA_PROCESO.xlsx;Extended Properties='Excel 12.0 Xml;HDR=Yes;'";


        //    using (OleDbConnection connection = new OleDbConnection(con))
        //    {
        //        connection.Open();
        //        OleDbCommand command = new OleDbCommand("select * from [DATA$]", connection);
        //        int cont = 0;
        //        using (OleDbDataReader dr = command.ExecuteReader())
        //        {
        //            while (dr.Read())
        //            {
        //                cont = cont + 1;
        //                row1Col0 = dr[0].ToString();
        //                row1Col1 = dr[1].ToString();
        //                row1Col2 = dr[2].ToString();
        //                row1Col3 = dr[3].ToString();
        //                row1Col4 = dr[4].ToString();
        //                row1Col5 = "2908070001210000";//dr[5].ToString();

        //                //MainFrameAdapter.SetTextOnScreen(15, 47, row1Col0);
        //                //MainFrameAdapter.SetTextOnScreen(17, 47, "9999");
        //                //MainFrameAdapter.SendKey(PcomKeys.PF2);
        //                Thread.Sleep(tiempo);
        //                //montoAbonado = MainFrameAdapter.GetScreenText(4, 4, 14);

        //                if (montoAbonado != null)
        //                {
        //                    Thread.Sleep(tiempo);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 25, row1Col1);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.FieldPlus);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);                             
        //                    //MainFrameAdapter.SetTextOnScreen(7, 32, row1Col2);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 35, row1Col3);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 39, row1Col4);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 43, row1Col5);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    ////MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP  _" + cont.ToString());
        //                    //MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP");
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF11);
        //                    //Thread.Sleep(tiempo);
        //                    //MainFrameAdapter.SetTextOnScreen(12, 19, ".");
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF3);


        //                    //string newColumn = "H";
        //                    //string newRow = "5";
        //                    //string worksheet2 = "DATA";

        //                    ////string sql2 = String.Format("UPDATE [{0}$] SET {1}{2}={3}", worksheet2, newColumn, newRow, "Hola");
        //                    //string commandString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1='{3}'", worksheet2, newColumn, newRow, 1);
        //                    //OleDbCommand objCmdSelect = new OleDbCommand(commandString, connection);
        //                    //objCmdSelect.ExecuteNonQuery();



        //                    //COMANDOS PARA EL SIGUIENTE REGISTRO
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF7);
        //                    //MainFrameAdapter.SetTextOnScreen(22, 7, "2");
        //                    //MainFrameAdapter.SendKey(PcomKeys.Enter);
        //                }
        //            }
        //        }
        //    }
        //}

        public void DoUIValidation()
        {
            throw new NotImplementedException();
        }

        public void ExtractElements()
        {
            throw new NotImplementedException();
        }
    }
}
