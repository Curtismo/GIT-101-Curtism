using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Net;

namespace TestImporter
{
	class TestReporting
	{
		public static ITestRun CreateTestRun(int testId)
		{
			NetworkCredential cred = new NetworkCredential("UserName", "Password");
			TfsTeamProjectCollection tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri("VSTSSiteBase"));

			tfs.Credentials = cred;
			tfs.Authenticate();
			tfs.EnsureAuthenticated();

			ITestManagementTeamProject project = tfs.GetService<ITestManagementService>().GetTeamProject("Schwans Company");

			// find the test case.
			ITestCase testCase = project.TestCases.Find(testId);
			string title = testCase.Title.ToString();

			// find test plan.
			int planId = Int32.Parse("testPlanId");
				//ConfigurationManager.AppSettings["TestPlanId"]);
			ITestPlan plan = project.TestPlans.Find(planId);

			// Create test configuration. You can reuse this instead of creating a new config everytime.
			ITestConfiguration config = CreateTestConfiguration(project, string.Format("My test config {0}", DateTime.Now));

			// Create test points. 
			IList<ITestPoint> testPoints = CreateTestPoints(project, plan, new List<ITestCase>() { testCase }, new IdAndName[] { new IdAndName(config.Id, config.Name) });

			// Create test run using test points.
			ITestRun run = CreateRun(project, plan, testPoints, title);
			return run;
		}
		public static void UpdateTestRun(ITestRun run, bool outcome, string comment)
		{
			run.Comment = comment;
			run.Save();
			// Query results from the run.
			ITestCaseResult result = run.QueryResults()[0];
			if (outcome == true)
			{
				// Fail the result.
				result.Outcome = TestOutcome.Failed;
			}
			else if (outcome == false)
			{
				// Pass the result.
				result.Outcome = TestOutcome.Passed;
			}
			result.State = TestResultState.Completed;
			result.Save();
		}
		private static ITestConfiguration CreateTestConfiguration(ITestManagementTeamProject project, string title)
		{
			ITestConfiguration configuration = project.TestConfigurations.Create();
			configuration.Name = title;
			configuration.Description = "DefaultConfig";
			configuration.Values.Add(new KeyValuePair<string, string>("Browser", "IE"));
			configuration.Save();
			return configuration;
		}
		public static IList<ITestPoint> CreateTestPoints(ITestManagementTeamProject project, ITestPlan testPlan, IList<ITestCase> testCases, IList<IdAndName> testConfigs)
		{
			IStaticTestSuite testSuite = CreateTestSuite(project);
			testPlan.RootSuite.Entries.Add(testSuite);
			testPlan.Save();
			testSuite.Entries.AddCases(testCases);
			testPlan.Save();
			testSuite.SetEntryConfigurations(testSuite.Entries, testConfigs);
			testPlan.Save();
			ITestPointCollection tpc = testPlan.QueryTestPoints("SELECT * FROM TestPoint WHERE SuiteId = " + testSuite.Id);
			return new List<ITestPoint>(tpc);
		}
		private static ITestRun CreateRun(ITestManagementTeamProject project, ITestPlan plan, IList<ITestPoint> points, string title)
		{
			ITestRun run = plan.CreateTestRun(false);
			foreach (ITestPoint tp in points)
			{
				run.AddTestPoint(tp, null);
				run.Title = title;
			}
			run.Save();
			return run;
		}
		private static IStaticTestSuite CreateTestSuite(ITestManagementTeamProject project)
		{
			// Create a static test suite.
			IStaticTestSuite testSuite = project.TestSuites.CreateStatic();
			testSuite.Title = "Static Suite";
			return testSuite;
		}
	}
}
