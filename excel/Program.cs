using System;
using NetOffice.ExcelApi;

namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Application ex = new Application();
            ex.Workbooks.Add();
            ex.Cells[1,1].Value = "Ford";
            ex.Cells[1,2].Value = "Fiesta";
            ex.Cells[1,3].Value = "1.8";
            ex.ActiveWorkbook.SaveAs(@"C:\Users\39694603870\Desktop\Projetos\projeto-5\excel\cliente.xlsx");
            ex.Quit();
        }
    }
}
