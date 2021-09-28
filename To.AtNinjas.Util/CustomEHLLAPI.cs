using EHLLAPI;
using System;
using System.Configuration;

namespace To.AtNinjas.Util
{
    public class CustomEHLLAPI
    {

        public void SendStr(string param)
        {
            String ReadyKey = ConfigurationManager.AppSettings["ReadKey"].ToString().Trim();
            if (ReadyKey.Equals("SI") == true)
            {
                Console.ReadKey();
            }
            var LogStr = param;
            switch (param)
            {
                case "@A": LogStr = "Alt"; break;
                case "@$": LogStr = "Alt Cursor"; break;
                case "@A@Q": LogStr = "AEhllapiFuncntion"; break;
                case "@<": LogStr = "Backspace"; break;
                case "@B": LogStr = "Backtab(Left Tab)"; break;
                case "@C": LogStr = "Clear"; break;
                case "@A@Y": LogStr = "Cmd(function) Key"; break;
                case "@V": LogStr = "Cursor Down"; break;
                case "@L": LogStr = "Cursor Left"; break;
                case "@Z": LogStr = "Cursor Right"; break;
                case "@A@J": LogStr = "Cursor Select"; break;
                case "@U": LogStr = "Cursor Up"; break;
                case "@D": LogStr = "Delete"; break;
                case "@S@x": LogStr = "Dup "; break;
                case "@q": LogStr = " End"; break;
                case "@E": LogStr = " Enter"; break;
                case "@F": LogStr = " Erase EOF"; break;
                case "@A@F": LogStr = " Erase Input"; break;
                case "@A@E": LogStr = " Field Exit"; break;
                case "@S@y": LogStr = " Field Mark"; break;
                case "@A@-": LogStr = " Field -"; break;
                case "@A@+": LogStr = " Field +"; break;
                case "@H": LogStr = " Help"; break;
                case "@A@X": LogStr = " Hexadecimal"; break;
                case "@0": LogStr = " Home"; break;
                case "@I": LogStr = " Insert"; break;
                case "@A@I": LogStr = " Insert Toggle"; break;
                case "@P": LogStr = " Local Print"; break;
                case "@N": LogStr = " New Line"; break;
                case "@u": LogStr = " Page Up"; break;
                case "@v": LogStr = " Page Down"; break;
                case "@A@t": LogStr = " Print(PC)"; break;
                case "@A@T": LogStr = " Print Screen"; break;
                case "@A@<": LogStr = " Record Backspace"; break;
                case "@R": LogStr = " Reset"; break;
                case "@S": LogStr = " Shift"; break;
                case "@A@H": LogStr = " Sys Request"; break;
                case "@T": LogStr = " Tab(Right Tab)"; break;
                case "@A@C": LogStr = " Test"; break;
                case "@x": LogStr = " PA1 "; break;
                case "@y": LogStr = " PA2 "; break;
                case "@z": LogStr = " PA3"; break;
                case "@+": LogStr = " PA4"; break;
                case "@%": LogStr = " PA5 "; break;
                case "@&": LogStr = " PA6 "; break;
                case "@'": LogStr = " PA7 "; break;
                case "@(": LogStr = " PA8 "; break;
                case "@)": LogStr = " PA9 "; break;
                case "@*": LogStr = " PA10 "; break;
                case "@1": LogStr = " PF1 / F1"; break;
                case "@2": LogStr = " PF2 / F2"; break;
                case "@3": LogStr = " PF3 / F3"; break;
                case "@4": LogStr = " PF4 / F4"; break;
                case "@5": LogStr = " PF5 / F5"; break;
                case "@6": LogStr = " PF6 / F6"; break;
                case "@7": LogStr = " PF7 / F7"; break;
                case "@8": LogStr = " PF8 / F8"; break;
                case "@9": LogStr = " PF9 / F9"; break;
                case "@a": LogStr = " PF10 / F10"; break;
                case "@b": LogStr = " PF11 / F11"; break;
                case "@c": LogStr = " PF12 / F12"; break;
                case "@d": LogStr = " PF13 / F13"; break;
                case "@e": LogStr = " PF14 / F14"; break;
                case "@f": LogStr = " PF15 / F15"; break;
                case "@g": LogStr = " PF16 / F16"; break;
                case "@h": LogStr = " PF17 / F17"; break;
                case "@i": LogStr = " PF18 / F18"; break;
                case "@j": LogStr = " PF19 / F19"; break;
                case "@k": LogStr = " PF20 / F20"; break;
                case "@l": LogStr = " PF21 / F21"; break;
                case "@m": LogStr = " PF22 / F22"; break;
                case "@n": LogStr = " PF23 / F23"; break;
                case "@o": LogStr = " PF24 / F24"; break;
                default: LogStr = param; break;
            }
            EhllapiWrapper.SendStr(param);
            Methods.LogProceso("Se envió el parámetro:\t " + param);
            EhllapiWrapper.Wait();
        }
        //public void SetCursorPos(int position)
        //{
        //    EhllapiWrapper.SetCursorPos(position);
        //    EhllapiWrapper.Wait();
        //}
        public void SetCursorPos(string position)
        {
            String ReadyKey = ConfigurationManager.AppSettings["ReadKey"].ToString().Trim();
            var Spliteo = position.Split(',');

            int x, y;
            Int32.TryParse(Spliteo[0], out x);
            Int32.TryParse(Spliteo[1], out y);

            EhllapiWrapper.SetCursorPos(Methods.GetWSPosition(x, y));
            Methods.LogProceso("Situado en la posición \t [" + x + "/" + y + "]");
            EhllapiWrapper.Wait();

            if (ReadyKey.Equals("SI") == true)
            {
                Console.ReadKey();
            }
        }
        //public string ReadScreen(int position, int lenght)
        //{
        //    var outParam= "";
        //    EhllapiWrapper.ReadScreen(position, lenght, out outParam);
        //    EhllapiWrapper.Wait();
        //    return outParam.Substring(0,lenght);
        //}

        public string ReadScreen(string position, int lenght)
        {
            var Spliteo = position.Split(',');
            int x, y;
            Int32.TryParse(Spliteo[0], out x);
            Int32.TryParse(Spliteo[1], out y);
            var outParam = "";
            EhllapiWrapper.ReadScreen(Methods.GetWSPosition(x, y), lenght, out outParam);
            Methods.LogProceso("Se leyó la posición \t\t [" + x + "/" + y + "]");
            EhllapiWrapper.Wait();
            return outParam.Substring(0, lenght);
        }

        public void Connect(string id)
        {
            EhllapiWrapper.Connect(id);
            EhllapiWrapper.Wait();
        }


    }
}
