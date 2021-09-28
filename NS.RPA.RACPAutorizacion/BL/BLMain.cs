using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using To.AtNinjas.Util;

namespace NS.RPA.RACPAutorizacion.BL
{
    class BLMain
    {
        CustomEHLLAPI ehllapi;
        private const int ROW_LENGTH = 80;
        public BLMain()
        {
            ehllapi = new CustomEHLLAPI();
            //ehllapi.Connect(SessionId);
        }

        public List<T> GetDataFromScreenList<T>(Func<string, T> formatMethod, int startYCoordenates,
         int pageSize, int rowsByItem, Func<string, int, string> readData/*MyDelegateType<T> formatMethod, */)
        {
            //string temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
            string temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);

            EhllapiWrapper.Wait();

            List<T> rows = new List<T>();
            int iItems = 0;

            if (temp.Trim() != "")
            {
                var obj = formatMethod(temp);

                if (obj != null)
                {
                    rows.Add(obj);
                    iItems = rowsByItem;
                }
            }

            int i = rowsByItem;
            while (temp.Trim() != "")
            {
                if (i >= pageSize)
                {
                    //ehllapi.SendStr("@v");

                    //TODO: si hay error en IBS ahí queda( deberíamos ver como obtener el valor de la x)
                    //temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    Methods.LogProceso("last " + temp);
                    Methods.LogProceso("last2 " + rows[rows.Count - iItems / rowsByItem].ToString());


                    if (temp == rows[rows.Count - iItems / rowsByItem].ToString()) { break; }
                    else
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                        }
                    }

                    i = rowsByItem;
                    iItems = rowsByItem;
                }
                else
                {
                    //TODO: depende de donde inicia la lista
                    //temp = ehllapi.ReadScreen(startYCoordenates + i + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData((startYCoordenates + i).ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    if (temp.Trim() != "")
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                            iItems += rowsByItem;
                        }
                    }
                    i += rowsByItem;
                }
            }

            return rows;
        }


        public List<T> SetDataFromScreenList<T>(Func<string, T> formatMethod, int startYCoordenates,
         int pageSize, int rowsByItem, Func<string, int, string> readData/*MyDelegateType<T> formatMethod, */)
        {
            //string temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
            string temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);

            EhllapiWrapper.Wait();

            List<T> rows = new List<T>();
            int iItems = 0;

            if (temp.Trim() != "")
            {
                var obj = formatMethod(temp);

                if (obj != null)
                {
                    rows.Add(obj);
                    iItems = rowsByItem;
                }
            }

            int i = rowsByItem;
            while (temp.Trim() != "")
            {
                if (i >= pageSize)
                {
                    //ehllapi.SendStr("@v");

                    //TODO: si hay error en IBS ahí queda( deberíamos ver como obtener el valor de la x)
                    //temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    Methods.LogProceso("last " + temp);
                    Methods.LogProceso("last2 " + rows[rows.Count - iItems / rowsByItem].ToString());


                    if (temp == rows[rows.Count - iItems / rowsByItem].ToString()) { break; }
                    else
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                        }
                    }

                    i = rowsByItem;
                    iItems = rowsByItem;
                }
                else
                {
                    //TODO: depende de donde inicia la lista
                    //temp = ehllapi.ReadScreen(startYCoordenates + i + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData((startYCoordenates + i).ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    if (temp.Trim() != "")
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                            iItems += rowsByItem;
                        }
                    }
                    i += rowsByItem;
                }
            }

            return rows;
        }

        public void SetApprovalPass(string pass)
        {
            ehllapi.SetCursorPos("3,14");
            EhllapiWrapper.Wait();
            ehllapi.SendStr(pass);
            EhllapiWrapper.Wait();
        }

        public string ReadScreen(string position, int lenght)
        {
            return ehllapi.ReadScreen(position + ",1", lenght);
        }



        public object GetLoanPaymentMovement(string plot)
        {
            //me aseguro de que no tendré un fuera de rango... .
            //plot = plot.PadRight(80 * rowsByItem);

            //TODO: manejo de nulos
            //if (plot.Substring(20, 10)

            /*
            " 13/12/07 13/12/07 ND        33,600.00                                          "
            " 13/12/07 13/12/07 RA APERTURA 010.500000                                       "
            " 12/06/19 12/06/19 RA TASA 008.500000 AL 008.000000                             "
            */

            string cd_Tr = plot.Substring(19, 2).Trim();
            decimal? principal = null;
            decimal? rate = null;
            string principalCR = null;
            if (cd_Tr == "RA")
            {
                if (plot.Contains("APERTURA"))
                {
                    rate = DecimalTryParse(plot.Substring(30, 11)) / 100;
                }
                else if (plot.Contains("TASA"))
                {
                    rate = DecimalTryParse(plot.Substring(40, 11)) / 100;
                }
            }
            else
            {
                principal = DecimalTryParse(plot.Substring(21, 17));
                principalCR = plot.Substring(38, 2);
            }

            return new object();
            //return new LoanPaymentMovement
            //{
            //    Plot = plot,
            //    //chapar hasta antes del inicio del sgte elemento
            //    ProcessDate = DatetimeTryParse(plot.Substring(1, 8)),
            //    ValueDate = DatetimeTryParse(plot.Substring(10, 8)),
            //    Cd_Tr = cd_Tr,
            //    Principal = principal,
            //    PrincipalCR = principalCR,
            //    Rate = rate,
            //    OffsettingInterest = DecimalTryParse(plot.Substring(40, 12)),
            //    OffsettingInterestCR = plot.Substring(52, 2),

            //    //faltan los sgtes
            //    OverdueCharge = DecimalTryParse(plot.Substring(54, 12)), //se sabe el indice de inicio falta saber la longitud
            //    OverdueChargeCR = plot.Substring(66, 2),

            //    InterestAdjusment = DecimalTryParse(plot.Substring(68, 10)),
            //    InterestAdjusmentCR = plot.Substring(78, 2),
            //    Batch = plot.Substring(85, 6).Trim(),
            //    Aprobo = plot.Substring(97, 12).Trim(),
            //    Description = plot.Substring(112, 32).Trim(),

            //    Origin = plot.Substring(150, 7).Trim(),

            //};
        }

        public static decimal? DecimalTryParse(string value)
        {
            decimal result;
            if (!decimal.TryParse(value, out result))
                return null;
            return result;
        }

        public static DateTime? DatetimeTryParse(string value)
        {
            DateTime result;
            if (!DateTime.TryParse(value, out result))
                return null;
            return result;
        }


    }
}
