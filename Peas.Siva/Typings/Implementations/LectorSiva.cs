using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Peas.Siva.Typings.Interfaces;
using Peas.Siva.Models;
using OfficeOpenXml;
using System.IO;

namespace Peas.Siva.Typings.Implementations
{
	public class LectorSiva : ILectorExcel<SivaReporte>
	{
		private readonly string _fileName = "siva.xlsx";
		public LectorSiva()
		{
			//Console.WriteLine(Directory.GetCurrentDirectory());
		}
		
		public IEnumerable<SivaReporte> LeerTodo()
		{
			var listaSiva = new List<SivaReporte>();
			using (var file = File.OpenRead(_fileName))
			{
				ExcelPackage excelPackage = new ExcelPackage(file);
				ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[1];

				var celdas = excelWorksheet.Cells;

				var folioColumna = celdas.Where( x => x.Text.Equals("FOLIO")).First().Start.Column;
				var cuentaColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("CUENTA")).First().Start.Column;
				var contratoColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("CONTRATO")).First().Start.Column;
				var dnColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("DN")).First().Start.Column;
				var clienteColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("CLIENTE")).First().Start.Column;
				var planColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("PLAN")).First().Start.Column;
				var equipoColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("EQUIPO")).First().Start.Column;
				var plazoColumna = excelWorksheet.Cells.Where(x => x.Text.Equals("PLAZO")).First().Start.Column;

				for (int i = excelWorksheet.Dimension.Start.Row + 1; i < excelWorksheet.Dimension.End.Row; i++)
				{
					listaSiva.Add(new SivaReporte()
					{
						Folio = celdas[i, folioColumna].Value.ToString(),
						Cliente = celdas[i, clienteColumna].Value.ToString(),
						Contrato = celdas[i, contratoColumna].Value.ToString(),
						Cuenta = celdas[i, cuentaColumna].Value.ToString(),
						DN = celdas[i,dnColumna].Value.ToString(),
						Equipo = celdas[i, equipoColumna].Value.ToString(),
						Plan = celdas[i, planColumna].Value.ToString(),
						Plazo = celdas[i, plazoColumna].Value.ToString()
					});
				}
			}

			return listaSiva;
		}
	}
}
