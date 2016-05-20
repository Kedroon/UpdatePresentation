using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Net;
using HtmlAgilityPack;
using System.Threading;

namespace UpdatePresentation
{
    class Program
    {
        

        static void Main(string[] args)
        {

            WebClient web = new WebClient();
            web.UseDefaultCredentials = true;
            web.Proxy.Credentials = CredentialCache.DefaultCredentials;
            string html="";
            web.Credentials = CredentialCache.DefaultCredentials;
            
            try
            {
                html = web.DownloadString("https://ptax.bcb.gov.br/ptax_internet/consultarTodasAsMoedas.do?method=consultaTodasMoedas");
            }
            catch (Exception  err)
            {

                Console.WriteLine(err.Message);
            }
            
            
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            //Console.WriteLine(html);
            
            HtmlNode dolarcompra = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""USD"")]//td[4]");
            HtmlNode dolarvenda = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""USD"")]//td[5]");
            Console.WriteLine(dolarcompra.InnerHtml);
            Console.WriteLine(dolarvenda.InnerHtml);

            HtmlNode eurocompra = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""EUR"")]//td[4]");
            HtmlNode eurovenda = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""EUR"")]//td[5]");
            Console.WriteLine(eurocompra.InnerHtml);
            Console.WriteLine(eurovenda.InnerHtml);

            HtmlNode yencompra = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""JPY"")]//td[4]");
            HtmlNode yenvenda = doc.DocumentNode.SelectSingleNode(@"//tr[contains(td[3],""JPY"")]//td[5]");
            Console.WriteLine(yencompra.InnerHtml);
            Console.WriteLine(yenvenda.InnerHtml);
            
            


            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            string workbookPath = @"C:\Users\migue\Desktop\Cotacoes.xlsx";
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            string currentSheet = "Planilha1";
            //MessageBox.Show(currentSheet);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            excelWorksheet.Cells[5, 2] = dolarcompra.InnerHtml.Replace(",",".");
            excelWorksheet.Cells[5, 3] = dolarvenda.InnerHtml.Replace(",", ".");

            excelWorksheet.Cells[6, 2] = eurocompra.InnerHtml.Replace(",", ".");
            excelWorksheet.Cells[6, 3] = eurovenda.InnerHtml.Replace(",", ".");

            excelWorksheet.Cells[7, 2] = yencompra.InnerHtml.Replace(",", ".");
            excelWorksheet.Cells[7, 3] = yenvenda.InnerHtml.Replace(",", ".");

            excelWorkbook.Save();
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            excelApp.Quit();

            string powerpointPath = @"C:\Users\migue\Desktop\Cotacoes.pptx";
            PowerPoint.Application ppApp = new PowerPoint.Application();
            ppApp.Visible = MsoTriState.msoTrue;
            PowerPoint.Presentations oPresSet = ppApp.Presentations;
            PowerPoint._Presentation oPres = oPresSet.Open(powerpointPath,
                    MsoTriState.msoFalse, MsoTriState.msoFalse,
                    MsoTriState.msoTrue);
            oPres.UpdateLinks();
            oPres.Save();
            oPres.SlideShowSettings.Run();


            Thread.Sleep(2000);
            updateexcel(1,2,3, 4,5,6);
            oPres.UpdateLinks();
            oPres.Save();
            Thread.Sleep(2000);
            updateexcel(7, 8, 9, 10, 11, 12);
            oPres.UpdateLinks();
            oPres.Save();
            Thread.Sleep(2000);
            updateexcel(1, 5, 4, 175, 134, 242);
            oPres.UpdateLinks();
            oPres.Save();
            Thread.Sleep(2000);
            updateexcel(134, 134, 242, 2142, 42, 424);
            oPres.UpdateLinks();
            oPres.Save();
            Console.Read();
            oPres.Close();
            Marshal.ReleaseComObject(oPres);
            Marshal.ReleaseComObject(oPresSet);
            //Marshal.ReleaseComObject(ppApp);
            ppApp.Quit();


        }

        public static void updateexcel(Decimal num1, Decimal num2, Decimal num3 , Decimal num4, Decimal num5 , Decimal num6) {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            string workbookPath = @"C:\Users\migue\Desktop\Cotacoes.xlsx";
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            string currentSheet = "Planilha1";
            //MessageBox.Show(currentSheet);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            excelWorksheet.Cells[5, 2] = num1;
            excelWorksheet.Cells[5, 3] = num2;

            excelWorksheet.Cells[6, 2] = num3;
            excelWorksheet.Cells[6, 3] = num4;

            excelWorksheet.Cells[7, 2] = num5;
            excelWorksheet.Cells[7, 3] = num6;

            excelWorkbook.Save();
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            //Marshal.ReleaseComObject(excelApp);
            excelApp.Quit();

        }


    }
}
