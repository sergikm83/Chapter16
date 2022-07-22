using System;
using System.Reflection;

namespace LateBindingWithDynamic
{
    class Program
    {
        static void Main(string[] args)
        {
            AddWithReflection();
            AddWithDynamic();
        }

        private static void AddWithReflection()
        {
            Assembly asm = Assembly.Load("MathLibrary");
            try
            {
                // Получить метаданные для типа SimpleMath.
                Type math = asm.GetType("MathLibrary.SimpleMath");
                // Создать объект SimpleMAth на лету.
                object obj = Activator.CreateInstance(math);
                // Получить информацию о методе Add().
                MethodInfo mi = math.GetMethod("Add");
                // Вызвать метод Add с параметрами.
                object[] args = { 10, 70 };
                Console.WriteLine("Results is: {0}", mi.Invoke(obj, args));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static void AddWithDynamic()
        {
            Assembly asm = Assembly.Load("MathLibrary");
            try
            {
                // Получить метаданные для типа SimpleMath.
                Type math = asm.GetType("MathLibrary.SimpleMath");
                // Создать объект SimpleMath на лету.
                dynamic obj = Activator.CreateInstance(math);
                // Обратите внимание, насколько легко теперь вызывать метод Add().
                Console.WriteLine("Results is: {0}", obj.Add(10, 70));
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
