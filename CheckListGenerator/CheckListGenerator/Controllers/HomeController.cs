using CheckListGenerator.Models;
using CheckListGenerator.ViewModels;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CheckListGenerator.Controllers
{
	public class HomeController : Controller
	{
		private readonly IWebHostEnvironment _appEnvironment;

		public HomeController(IWebHostEnvironment appEnvironment)
		{
			_appEnvironment = appEnvironment;
		}

		public IActionResult Index()
		{
			return View();
		}



		[HttpPost]
		[AutoValidateAntiforgeryToken]
		public IActionResult GenerateCheckList(IndexViewModel model)
		{
			if (ModelState.IsValid)
			{
				bool eng = model.Language;
				bool isBugReport = false, isPRD = false;

				if (model.File == null || model.File.Length == 0)
				{
					ModelState.AddModelError(string.Empty, "Файл не завантажено або він порожній.");
					return RedirectToAction(actionName: nameof(Index), routeValues: model);
				}


				string[] documentWords = {
											"технічне завдання", "technical specification",
											"product requirements document", "документ вимог до продукту",
											"тз", "prd",
											"мета", "goal",
											"сценарії використання", "use cases",
											"функціональні вимоги", "functional requirements",
											"технічні вимоги", "technical requirements",
											"інтерфейси користувача", "user interfaces",
											"бюджет", "budget",
											"план розробки", "development plan",
											"графік реалізації", "implementation schedule",
											"ролі користувачів", "user roles",
											"архітектура системи", "system architecture",
											"вимоги безпеки", "security requirements",
											"тестування", "testing",
											"критерії прийняття", "acceptance criteria",
											"обмеження", "constraints",
											"потреби стейкхолдерів", "stakeholder needs",
											"середовище використання", "operating environment",
											"регуляторні вимоги", "regulatory requirements",
											"вимоги до надійності", "reliability requirements",
											"версіонування", "versioning"
										};

				string[] bugReportWords = {		"bug", "помилка",
												"issue", "проблема",
												"error", "помилковий",
												"fault", "несправність",
												"defect", "дефект",
												"problem", "неполадка",
												"failure", "збій",
												"malfunction", "неправильна робота",
												"glitch", "збій",
												"reproduce", "відтворення",
												"steps to reproduce", "кроки для відтворення",
												"expected result", "очікуваний результат",
												"actual result", "фактичний результат",
												"severity", "серйозність",
												"priority", "пріоритет",
												"status", "статус",
												"fix", "виправлення",
												"resolved", "вирішено",
												"workaround", "обхідний шлях",
												"patch", "патч",
												"log", "журнал",
												"error code", "код помилки",
												"crash", "аварійне завершення",
												"hang", "зависання",
												"slow performance", "повільна робота",
												"memory leak", "витік пам'яті",
												"compatibility issue", "проблема сумісності",
												"security flaw", "безпековий недолік"
											};


				XLWorkbook workbook;

				using (var stream = model.File.OpenReadStream())
				{
					using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, false))
					{
						var bodyText = doc.MainDocumentPart.Document.Body.InnerText;


						foreach (var item in documentWords)
						{
							if (bodyText.Contains(item, StringComparison.CurrentCultureIgnoreCase))
							{
								isPRD = true;
								break;
							}
						}
						foreach (var item in bugReportWords)
						{
							if (bodyText.Contains(item, StringComparison.CurrentCultureIgnoreCase))
							{
								isBugReport = true;
								break;
							}
						}						
					}
				}

				string path = "";
				string name = "";

				if (isPRD && !isBugReport)
				{
					if (eng)
					{
						path = _appEnvironment.WebRootPath + @"\checklists\CheckList (ТЗ).xlsx";
						name = "CheckList for PRD (Product requirements document)";
					}
					else
					{
						path = _appEnvironment.WebRootPath + @"\checklists\Чек-ліст (ТЗ).xlsx";
						name = "Чек-ліст для ТЗ (Технічне завдання)";
					}

                    workbook = new XLWorkbook(path);

                    DocumentViewModel dvm = new DocumentViewModel() { Workbook = workbook, Path = path, FileName = name };

                    return View(dvm);
                }
				else if (!isPRD && isBugReport)
				{
					if (eng)
					{
						path = _appEnvironment.WebRootPath + @"\checklists\CheckList (bug report).xlsx";
						name = "CheckList for Bug report";
					}
					else
					{
						path = _appEnvironment.WebRootPath + @"\checklists\Чек-ліст (bug report).xlsx";
						name = "Чек-ліст для звіту про помилку";
					}

					workbook = new XLWorkbook(path);

					DocumentViewModel dvm = new DocumentViewModel() { Workbook = workbook, Path = path, FileName = name };

                    return View(dvm);
				}
				else
				{
					ModelState.AddModelError(string.Empty, "Неможливо визначити тип файлу.");
				}
			}
			ModelState.AddModelError(string.Empty, "Файл не завантажено або він порожній.");
			return RedirectToAction(actionName: nameof(Index), routeValues: model);
		}

		public IActionResult Download(string path)
		{
			string filePath = path;
            string fileName = "CheckList.xlsx";

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }


		public IActionResult Privacy()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}
	}
}