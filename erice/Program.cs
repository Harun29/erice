using System;

namespace erice
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Unesite vrijednost chol:");
            double chol = double.Parse(Console.ReadLine());

            Console.WriteLine("Unesite vrijednost age:");
            int age = int.Parse(Console.ReadLine());

            Console.WriteLine("Unesite vrijednost sbp:");
            double sbp = double.Parse(Console.ReadLine());

            Console.WriteLine("Da li je pacijent dijabetičar? (true/false):");
            bool diabetic = bool.Parse(Console.ReadLine());

            Console.WriteLine("Da li je pacijent pušač? (true/false):");
            bool smoker = bool.Parse(Console.ReadLine());


            int row = 0;
            int col = 0;
            ExcelReader excel = null;
            int start = 4;
            int finish = 24;
            int reading = 1;

            try
            {
                excel = new ExcelReader(@"C:\Users\Korisnik\Desktop\c#\erice\erice\Erice2.xlsx", 1);
                bool breakCheck = false;
                while (!breakCheck)
                {
                    if (row == 0)
                    {
                        for (int i = start; i <= finish; i++)
                        {
                            string cellValue = excel.ReadCell(i, reading);

                            if (int.TryParse(cellValue, out int cellIntValue))
                            {
                                if ((age > cellIntValue) && (cellIntValue != 0) && (reading == 1))
                                {
                                    start = i - 3;
                                    finish = i;
                                    reading = 3;
                                    break;
                                }
                                else if ((sbp > cellIntValue) && reading == 3)
                                {
                                    row = i;
                                    start = 4;
                                    finish = 17;
                                    reading = 1;
                                    break;
                                }
                            }
                            else
                            {
                                Console.WriteLine("Failed to parse cell value to integer.");
                            }
                        }
                    }
                    else
                    {
                        for (int i = start; i <= finish; i++)
                        {
                            string cellValue = excel.ReadCell(reading, i);

                            if (diabetic && (cellValue == "Diabetics") && (reading == 1))
                            {
                                start = 4;
                                finish = 11;
                                reading = 2;
                                break;
                            }
                            else if (!diabetic && (cellValue == "Non diabetics") && (reading == 1))
                            {
                                start = 12;
                                finish = 19;
                                reading = 2;
                                break;
                            }
                            else if (smoker && (cellValue == "Smokers") && (reading == 2))
                            {
                                start += 4;
                                reading = 3;
                                break;
                            }
                            else if (!smoker && (cellValue == "Non smokers") && (reading == 2))
                            {
                                finish -= 4;
                                reading = 3;
                                break;
                            }

                            else if (double.TryParse(cellValue, out double cellIntValue))
                            {
                                if ((chol < cellIntValue) && reading == 3)
                                {
                                    col = i;
                                    breakCheck = true;
                                    break;
                                }
                                else if ((chol > 7.8) && reading == 3)
                                {
                                    col = finish;
                                    breakCheck = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                string result = excel.ReadCell(row, col);
                Console.WriteLine(result);
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
