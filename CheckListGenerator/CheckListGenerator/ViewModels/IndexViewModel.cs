using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace CheckListGenerator.ViewModels
{
	public class IndexViewModel
	{
		[Key]
		public Guid? Id { get; set; }

		[Required]
		[Display(Name = "Language")]
		public bool Language { get; set; }


		[Display(Name = "File")]
		public IFormFile? File { get; set; }
		public string? FileName { get; set; }
		public string? FileFullName { get; set; }

	}
}
