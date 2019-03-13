using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Data;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.IO;
using System.Text;

namespace TestImporter
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new Form1());
		}
		public static Project GetTeamProject(string uri, string name)
		{
			TfsTeamProjectCollection tfs;
			NetworkCredential cred = new NetworkCredential("User", "password");
			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri));

			//tfs.Credentials = cred;
			//tfs.Authenticate();
			tfs.EnsureAuthenticated();

			var workItemStore = new WorkItemStore(tfs);

			var project = (from Project pr in workItemStore.Projects
						   where pr.Name == name
						   select pr).FirstOrDefault();
			if (project == null)
				throw new Exception($"Unable to find {name} in {uri}");

			return project;
		}
		public static int AddsharedSteps(SharedStepsObject sharedStepsObject, int newSharedStepId)
		{
			TfsTeamProjectCollection tfs;

			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(sharedStepsObject.uri)); // https://mytfs.visualstudio.com/DefaultCollection
			tfs.Authenticate();
			ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
			ITestManagementTeamProject testProject = service.GetTeamProject(sharedStepsObject.project);
			ISharedStep sharedStep = testProject.SharedSteps.Find(newSharedStepId);
			for (int i = 0; i < sharedStepsObject.actionSteps.Count(); i++)
			{
				ITestStep newStep = sharedStep.CreateTestStep();
				newStep.Title = sharedStepsObject.actionSteps[i];
				newStep.ExpectedResult = sharedStepsObject.results[i];
				sharedStep.Actions.Add(newStep);
			}
			sharedStep.Save();
			return sharedStep.Id;
		}
		public static int CreateSharedStep(SharedStepsObject sharedStepsObject)
		{
			WorkItemType workItemType = sharedStepsObject.project.WorkItemTypes["Shared Steps"];
			WorkItem newSharedStep = new WorkItem(workItemType)
			{
				Title = sharedStepsObject.title,
				AreaPath = sharedStepsObject.testedWorkItem.AreaPath,
				IterationPath = sharedStepsObject.testedWorkItem.IterationPath
			};
			ActionResult result = Program.CheckValidationResult(newSharedStep);
			if (result.Success)
			{
				Program.AddsharedSteps(sharedStepsObject, newSharedStep.Id);
			}
			TfsTeamProjectCollection tfs;
			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri((sharedStepsObject.uri)));
			tfs.Authenticate();
			ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
			ITestManagementTeamProject testProject = service.GetTeamProject(sharedStepsObject.project);
			return newSharedStep.Id;
		}
		public static void AddTestCaseSteps(SharedStepsObject sharedStepsObject, int testCaseId)
		{
			TfsTeamProjectCollection tfs;
			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(sharedStepsObject.uri));
			tfs.Authenticate();
			ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
			ITestManagementTeamProject testProject = service.GetTeamProject(sharedStepsObject.project);
			ITestCase testCase = testProject.TestCases.Find(testCaseId);
			ISharedStepReference sharedStepReference = testCase.CreateSharedStepReference();
			int sharedStepId = sharedStepsObject.id;
			sharedStepReference.SharedStepId = sharedStepId;
			testCase.Actions.Add(sharedStepReference);
			if (sharedStepsObject.automationValue.ToString() != "")
			{
				testCase.CustomFields["Automation Status1"].Value = sharedStepsObject.automationValue.ToString();//"Planned";
			}

			testCase.Save();
		}
		public static ActionResult CreateNewTestCase(SharedStepsObject sharedStepsObject, int parentId)
		{
			WorkItemType workItemType = sharedStepsObject.project.WorkItemTypes["Test Case"];
			WorkItem newTestCase = new WorkItem(workItemType)
			{
				Title = sharedStepsObject.title,
				Description = sharedStepsObject.description,
				AreaPath = sharedStepsObject.testedWorkItem.AreaPath,
				IterationPath = sharedStepsObject.testedWorkItem.IterationPath
			};
			ActionResult result = CheckValidationResult(newTestCase);
			if (result.Success)
			{
				CreateTestedByLink(sharedStepsObject, parentId);
				AddTestCaseSteps(sharedStepsObject, newTestCase.Id);
			}

			return result;
		}
		public static ActionResult CheckValidationResult(WorkItem workItem)
		{
			var validationResult = workItem.Validate();
			ActionResult result = null;
			if (validationResult.Count == 0 || validationResult.Count == 1)
			{
				workItem.Save();
				result = new ActionResult()
				{
					Success = true,
					Id = workItem.Id
				};
			}
			return result;
		}
		private static void CreateTestedByLink(SharedStepsObject sharedStepsObject, int parentId)
		{
			WorkItemStore workItemStore = sharedStepsObject.project.Store;
			Project teamProject = workItemStore.Projects["Schwans Company"];
			var linkTypes = workItemStore.WorkItemLinkTypes;
			WorkItemLinkType testedBy = linkTypes.FirstOrDefault(lt => lt.ForwardEnd.Name == "Tested By");
			WorkItemLinkTypeEnd linkTypeEnd = testedBy.ForwardEnd;
			sharedStepsObject.testedWorkItem.Links.Add(new RelatedLink(linkTypeEnd, sharedStepsObject.id + 1));
			var result = CheckValidationResult(sharedStepsObject.testedWorkItem);
		}
		public static WorkItem GetWorkItem(string uri, int testedWorkItemId)
		{
			TfsTeamProjectCollection tfs;
			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
			tfs.Authenticate();
			var workItemStore = new WorkItemStore(tfs);
			WorkItem workItem = workItemStore.GetWorkItem(testedWorkItemId);
			workItem.History.Insert(0, "");
			return workItem;
		}
		internal static void CreateTestsWithSharedSteps(string tFSUri, Microsoft.Office.Interop.Excel.Application excel, string path, TextBox textBox1, TextBox URL, TextBox projectName)
		{
			tFSUri = "https://" + URL.Text;
			int currentRow = 2;
			while (true)
			{
				int parentId = ExcelWork.GetParentId(currentRow, excel, path);
				if (parentId == 0)
				{
					textBox1.AppendText("\r\n No parent item found, ending import \r\n");
					break;
				}
				SharedStepsObject sharedStepObject = new SharedStepsObject
				{
					uri = "https://" + URL.Text,
					project = Program.GetTeamProject(tFSUri, projectName.Text),
					testedWorkItem = Program.GetWorkItem(tFSUri, parentId),
					title = ExcelWork.GetTitle(currentRow, excel, path),
					description = ExcelWork.GetDescrition(currentRow, excel, path),
					paramExcelFile = ExcelWork.GetParamFileName(currentRow, excel, path),
					automationValue = ExcelWork.GetAutomationValue(currentRow, excel, path),
					paramSheet = ExcelWork.GetParamSheetName(currentRow, excel, path)
				};
				currentRow++;
				sharedStepObject.actionSteps = ExcelWork.GetSteps(currentRow, excel, path).ToArray();
				sharedStepObject.results = ExcelWork.GetResults(currentRow, sharedStepObject.actionSteps.Count(), excel, path).ToArray();
				currentRow = currentRow + sharedStepObject.actionSteps.Count();
				if (sharedStepObject.paramExcelFile != "")
				{
					CreateIterativeCases(sharedStepObject, excel, textBox1, parentId);
				}
				else
				{
					textBox1.AppendText("\r\n Case started for " + sharedStepObject.title);
					sharedStepObject.id = Program.CreateSharedStep(sharedStepObject);
					var resultTestCase = Program.CreateNewTestCase(sharedStepObject, parentId);
					textBox1.AppendText("\r\n Case created for " + sharedStepObject.title + "\r\n");
				}
			}
		}
		public static void CreateIterativeCases(SharedStepsObject sharedStepObject, Microsoft.Office.Interop.Excel.Application excel, TextBox textBox1, int parentId)
		{

			string paramPath = Form1.paramPath.ToString() + sharedStepObject.paramExcelFile;
			sharedStepObject.paramTable = ExcelWork.GetParams(excel, paramPath, sharedStepObject.paramSheet.ToString());
			string titleHold = sharedStepObject.title;
			int NumOfActions = sharedStepObject.actionSteps.Count();
			string[] holdSteps = new string[NumOfActions];
			for (int i = 0; i < NumOfActions; i++)
			{
				holdSteps[i] = sharedStepObject.actionSteps[i];
			}
			string[] holdResults = new string[NumOfActions];
			for (int i = 0; i < NumOfActions; i++)
			{
				holdResults[i] = sharedStepObject.results[i];
			}

			for (int iteration = 0; iteration < sharedStepObject.paramTable.Rows.Count; iteration++)
			{
				sharedStepObject.title = titleHold + " - " + (iteration + 1).ToString();
				textBox1.AppendText("\r\n Case started for " + sharedStepObject.title);
				sharedStepObject.actionSteps = ReturnStepsReplaceParams(textBox1, sharedStepObject.actionSteps, sharedStepObject, iteration);
				sharedStepObject.results = ReturnStepsReplaceParams(textBox1, sharedStepObject.results, sharedStepObject, iteration);
				sharedStepObject.id = Program.CreateSharedStep(sharedStepObject);
				var resultTestCase = Program.CreateNewTestCase(sharedStepObject, parentId);
				textBox1.AppendText("\r\n Case created for " + sharedStepObject.title + "\r\n");
				//reset steps and results
				for (int i = 0; i < NumOfActions; i++)
				{
					sharedStepObject.actionSteps[i] = holdSteps[i];
				}
				for (int i = 0; i < NumOfActions; i++)
				{
					sharedStepObject.results[i] = holdResults[i];
				}
			}
		}

		public static string[] ReturnStepsReplaceParams(TextBox textBox1, string[] steps, SharedStepsObject sharedStepObject, int iteration)
		{
			int stepNum = 0;
			foreach (string step in steps)
			{
				while (step.Contains("${"))
				{
					foreach (DataColumn column in sharedStepObject.paramTable.Columns)
					{
						if (step.Contains("${" + column.ColumnName + "}"))
						{
							try
							{
								steps[stepNum] = step.Replace("${" + column.ColumnName + "}", sharedStepObject.paramTable.Rows[iteration][column.ColumnName].ToString());

							}
							catch (Exception e)
							{
								textBox1.AppendText("\r\n Exception found while creating " + sharedStepObject.title + " : \r\n" + e.Message.ToString());
							}
						}
					}
				}
				stepNum++;
			}
			return steps;
		}

		public static WorkItem CreateE2ETestCase(SharedStepsObject sharedStepsObject, int parentId)
		{
			WorkItemType workItemType = sharedStepsObject.project.WorkItemTypes["Test Case"];
			WorkItem newTestCase = new WorkItem(workItemType)
			{
				Title = sharedStepsObject.title,
			};
			ActionResult result = CheckValidationResult(newTestCase);
			if (result.Success)
			{
				AddE2ESteps(sharedStepsObject, result.Id);
				//CreateTestedByLink(sharedStepsObject, parentId);
			}

			return newTestCase;
		}

		public static void AddE2ESteps(SharedStepsObject sharedStepsObject, int testCaseId)
		{
			TfsTeamProjectCollection tfs;

			tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(sharedStepsObject.uri));
			tfs.Authenticate();
			ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
			ITestManagementTeamProject testProject = service.GetTeamProject(sharedStepsObject.project);

			ITestCase testCase = testProject.TestCases.Find(testCaseId);
			foreach (string step in sharedStepsObject.actionSteps)
			{
				int sharedStepId = GetsharedStepId(step, sharedStepsObject);
				ISharedStepReference sharedStepReference = testCase.CreateSharedStepReference();
				sharedStepReference.SharedStepId = sharedStepId;
				testCase.Actions.Add(sharedStepReference);
			}
			testCase.Save();

		}

		public static int GetsharedStepId(string step, SharedStepsObject sharedStepsObject)
		{
			WorkItemStore workItemStore = sharedStepsObject.project.Store;
			Project teamProject = workItemStore.Projects["Schwans Company"];
			WorkItemCollection workItemCollection = workItemStore.Query(
							" SELECT [System.Id], [System.WorkItemType]," +
							" [System.Title] " +
							" FROM WorkItems " +
							" WHERE [System.TeamProject] = '" + teamProject.Name +
							"'  and [System.Title] = '" + step +
							"' and [System.WorkItemType] != 'Test Case" +
							"' ORDER BY [System.Id]");

			int SharedStepId = workItemCollection[0].Id;
			return SharedStepId;
		}

		public static string[] GetSharedStepsForE2E(Microsoft.Office.Interop.Excel.Application excel, SharedStepsObject sharedStepsObject)
		{
			ExcelWork.GetStepsForE2E(1, excel, sharedStepsObject.paramSheet);
			return sharedStepsObject.actionSteps;
		}
	}
}

