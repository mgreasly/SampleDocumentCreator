using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleDocumentCreator
{
    internal class LinkGenerator
    {
        public static string RandomLink()
        {
            var rnd = new Random();
            var i = rnd.Next(4);
            switch (i)
            {
                case 0: return BrokenHyperLink();
                case 1: return WorkingHyperLink();
                case 2: return ContentManagerLink();
                case 3: return ContentManagerLink();
                default:
                    Console.WriteLine($"RandomLink {i}");
                    return WorkingHyperLink();
            }
        }

        public static string BrokenHyperLink()
        {
            return $"http://{Guid.NewGuid()}.com.au";
        }

        public static string WorkingHyperLink()
        {
            var rnd = new Random();
            switch (rnd.Next(7))
            {
                case 0:
                    return "https://reqres.in/";
                case 1:
                    return "https://httpbin.org/";
                case 2:
                    return "http://dummy.restapiexample.com/";
                case 3:
                    return "https://jsonplaceholder.typicode.com/";
                case 4:
                    return "https://fakerestapi.azurewebsites.net/";
                case 5:
                    return "https://www.programmableweb.com/apis/directory";
                case 6:
                    return "https://developers.google.com/maps/documentation";
                case 7:
                    return "https://developer.github.com/v3/repos/";
                default:
                    return "http://www.microsoft.com";
            }
        }

        public static string ContentManagerLink()
        {
            var rnd = new Random();

            var id1 = rnd.Next(10, 99);
            var id2 = rnd.Next(1000, 9999);
            return $"trim://D-{id1}-{id2}/?db=PR&view";
        }
    }
}