using System.Collections.Generic;

namespace TestImporter
{
	public class ActionResult
	{
		public bool Success { get; set; }
		public List<string> ErrorCodes { get; set; }
		public int Id { get; set; }
	}
}