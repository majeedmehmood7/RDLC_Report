using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.IO;
using AspNetCore.Reporting;
using Microsoft.Extensions.Configuration;

namespace Rajby_web.Controllers
{
	public class HomeController : Controller
	{
		private readonly IWebHostEnvironment _webHostEnvironment;
		private readonly string _connectionString;

		public HomeController(IWebHostEnvironment webHostEnvironment, IConfiguration configuration)
		{
			_webHostEnvironment = webHostEnvironment;
			_connectionString = configuration.GetConnectionString("RajbyDev");
		}

		public IActionResult Index()
		{
			return View();
		}

		public IActionResult GeneratePDF(int costingId)
		{
			try
			{
				var dataTable = GetCostingData(costingId);

				if (dataTable.Rows.Count == 0)
				{
					return NotFound("No data found for the given Costing ID.");
				}

				string mimetype = "";
				int extension = 1;

				var reportPath = Path.Combine(_webHostEnvironment.WebRootPath, "Reports", "rptPreCostingFull.rdlc");

				if (!System.IO.File.Exists(reportPath))
				{
					return NotFound("The report definition file was not found.");
				}

				Dictionary<string, string> parameters = new Dictionary<string, string>
				{
					{ "CostingId", costingId.ToString() }
				};

				LocalReport localReport = new LocalReport(reportPath);
				localReport.AddDataSource("dsPreCosting", dataTable);

				var result = localReport.Execute(RenderType.Pdf, extension, parameters, mimetype);

				return File(result.MainStream, "application/pdf", $"CostingReport_{costingId}.pdf");
			}
			catch (Exception ex)
			{
				return StatusCode(500, $"Internal server error: {ex.Message}");
			}
		}

		public IActionResult ExportToExcel(int costingId)
		{
			try
			{
				var dataTable = GetCostingData(costingId);

				if (dataTable.Rows.Count == 0)
				{
					return NotFound("No data found for the given Costing ID.");
				}

				string mimetype = "";
				int extension = 1;

				var reportPath = Path.Combine(_webHostEnvironment.WebRootPath, "Reports", "rptPreCostingFull.rdlc");

				if (!System.IO.File.Exists(reportPath))
				{
					return NotFound("The report definition file was not found.");
				}

				Dictionary<string, string> parameters = new Dictionary<string, string>
				{
					{ "CostingId", costingId.ToString() }
				};

				LocalReport localReport = new LocalReport(reportPath);
				localReport.AddDataSource("dsPreCosting", dataTable);

				var result = localReport.Execute(RenderType.Excel, extension, parameters, mimetype);

				return File(result.MainStream, "application/vnd.ms-excel", $"CostingReport_{costingId}.xls");
			}
			catch (Exception ex)
			{
				return StatusCode(500, $"Internal server error: {ex.Message}");
			}
		}

		private DataTable GetCostingData(int costingId)
		{
			var dataTable = new DataTable();

			try
			{
				using (var connection = new SqlConnection(_connectionString))
				{
					string query = @"
                        SELECT 
                            * 
                        FROM vcmsPreCosting
                        WHERE costingId = @CostingId";

					using (var command = new SqlCommand(query, connection))
					{
						command.Parameters.AddWithValue("@CostingId", costingId);

						var adapter = new SqlDataAdapter(command);
						adapter.Fill(dataTable);
					}
				}
			}
			catch (SqlException sqlEx)
			{
				throw new Exception("SQL error occurred while retrieving data: " + sqlEx.Message, sqlEx);
			}
			catch (Exception ex)
			{
				throw new Exception("An error occurred while retrieving data.", ex);
			}

			return dataTable;
		}
	}
}
