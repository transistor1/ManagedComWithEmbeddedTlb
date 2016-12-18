using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComLib
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ComWithEmbeddedTypeLib : IComWithEmbeddedTypeLib
    {
        [DllExport("CreateObject", CallingConvention = CallingConvention.StdCall)]
        public static IComWithEmbeddedTypeLib CreateObject()
        {
            return new ComWithEmbeddedTypeLib();
        }
        public string HelloWorld(string name)
        {
            return String.Format("Hello, {0}!", name);
        }
    }
}
