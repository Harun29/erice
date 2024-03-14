using System;

namespace erice
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReader excel = null;
            try
            {
                excel = new ExcelReader(@"C:\Users\Korisnik\Desktop\c#\erice\erice\UKUPNI_BODOVI.xls", 1);
                for (int i = 1; i <= 10; i++)
                {
                    Console.WriteLine(excel.ReadCell(i, 2));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                if (excel != null)
                {
                    excel.Close();
                }
            }
        }
    }
}
