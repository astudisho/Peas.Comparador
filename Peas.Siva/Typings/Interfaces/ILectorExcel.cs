using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Peas.Siva.Typings.Interfaces
{
	public interface ILectorExcel<T> where T : class
	{
		IEnumerable<T> LeerTodo();

	}
}
