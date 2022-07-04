using DocentricConsoleApplicationNetCore;
using System;

namespace DocentricConsoleApplication
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ImageGeneration imageGeneration = new();
            string? file = imageGeneration.GenerateReport();
            Console.WriteLine(file);
        }
    }
}