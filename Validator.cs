using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Validator
    {
        public int check_int(string x)
        {
            int i = 0;
            int res = 0;
            bool flag = true;
            while (flag)
            {
                if (int.TryParse(x, out i) == false)
                {
                    Console.WriteLine("Не удалось распознать число типа int, попробуйте еще раз!");
                    Console.Write("Введите число: ");
                    x = Console.ReadLine();
                }
                else
                {
                    res = Convert.ToInt32(x);
                    flag = false;
                }
            }
            return res;
        }

        public double check_double(string x)
        {
            double i = 0;
            double res = 0;
            bool flag = true;
            while (flag)
            {
                if (double.TryParse(x, out i) == false)
                {
                    Console.WriteLine("Не удалось распознать число типа double, попробуйте еще раз!");
                    Console.Write("Введите число: ");
                    x = Console.ReadLine();
                }
                else
                {
                    res = Convert.ToDouble(x);
                    flag = false;
                }
            }
            return res;
        }
    }
}
