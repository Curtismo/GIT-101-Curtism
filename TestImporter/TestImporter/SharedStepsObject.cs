using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Data;

namespace TestImporter
{
	internal class SharedStepsObject
	{
		internal string[] actionSteps;
		internal string[] results;
		internal string automationValue;
		internal string description;
		internal string paramExcelFile;
		internal object paramSheet;
		internal Project project;
		internal WorkItem testedWorkItem;
		internal string title;
		internal string uri;
		internal int id;
		internal DataTable paramTable;
	}
}