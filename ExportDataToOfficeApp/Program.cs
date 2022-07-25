using System;
using System.Collections.Generic;

namespace ExportDataToOfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Car> carsInStock = new List<Car>
            {
                new Car{ Color = "Green", Make = "VW", PetName = "Mary" },
                new Car{ Color = "Red", Make = "Saab", PetName = "Mel" },
                new Car{ Color = "Black", Make = "Ford", PetName = "Hank" },
                new Car{ Color = "Yellow", Make = "BMW", PetName = "Davie" }
            };
        }
    }
}
