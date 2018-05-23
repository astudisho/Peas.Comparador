using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Peas.Siva.Typings.Implementations;

namespace Peas.Comparador
{
	class Program
	{
		static void Main(string[] args)
		{
			var lectorSiva = new LectorSiva();

			lectorSiva.LeerTodo();



			Console.ReadLine();
		}
	}
}
