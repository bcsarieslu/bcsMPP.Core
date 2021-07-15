using NUnit.Framework;
using bcsMPP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;

namespace bcsMPP.Core.Tests
{
    [TestFixture()]
    public class MPPTests
    {
        [Test()]
        public void getProcessPlanStructureTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item requestItem = inn.newItem("mpp_ProcessPlan", "mpp_getProcessPlanStructure");
            requestItem.setProperty("plan_ids", "");
            requestItem.setProperty("with_details_ids", "007519494C63433EBA748D90F54B7DD5");
            requestItem.setProperty("lang_code", "");
            requestItem.setProperty("bcs_location", "");

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.getProcessPlanStructure(requestItem);

            connection.Logout();

            Assert.Fail();
        }

        [Test()]
        public void GetTestWorkHourTreeGridTest()
        {
            HttpServerConnection connection = GetServerConnection();

            Innovator inn = IomFactory.CreateInnovator(connection);

            Item requestItem = inn.newItem("mpp_ProcessPlan");
            requestItem.setProperty("processplan_id", "007519494C63433EBA748D90F54B7DD5");
            requestItem.setProperty("processplan_number", "11111");
            requestItem.setProperty("processplan_name", "22222");
            requestItem.setProperty("location_id", "");

            bcsMPP.Core.MPP bcsMPP = new bcsMPP.Core.MPP(inn);
            Item result = bcsMPP.GetTestWorkHourTreeGrid(requestItem);

            connection.Logout();

            Assert.Pass("");
        }


        public HttpServerConnection GetServerConnection()
        {
            HttpServerConnection connection = IomFactory.CreateHttpServerConnection("http://localhost/arasmpp12sp9", "MPP12SP9", "admin", "innovator");
            Item loginResult = connection.Login();
            if (loginResult.isError())
            {
                Assert.Fail("登录失败");
            }
            return connection;
        }
    }
}