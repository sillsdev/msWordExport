// ---------------------------------------------------------------------------------------------
#region // Copyright (c) 2013, SIL International. All Rights Reserved.
// <copyright from='2013' to='2013' company='SIL International'>
//		Copyright (c) 2013, SIL International. All Rights Reserved.
//
//		Distributable under the terms of either the Common Public License or the
//		GNU Lesser General Public License, as specified in the LICENSING.txt file.
// </copyright>
#endregion
//
// File: ProgramTest.cs
// Responsibility: Greg Trihus
// License: LGPL2
// ---------------------------------------------------------------------------------------------
using System.IO;
using Test;
using msWordExport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Xsl;

namespace TestProject
{
    
    
    /// <summary>
    ///This is a test class for ProgramTest and is intended
    ///to contain all ProgramTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ProgramTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        static TestFiles _tf = null;
        //Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
            _tf = new TestFiles("TestProject");
        }
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for ApplyAfterRules
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void ApplyAfterRulesTest()
        {
            const string TestDataName = "after.xhtml";
            const string ClassName = "xsensenumber";
            var afterRules = new Dictionary<string, string>(){{ClassName, ") "}};
            var xDoc = new XmlDocument();
            xDoc.Load(_tf.Input(TestDataName));
            XmlNode node = xDoc.DocumentElement.SelectSingleNode("//*[contains(@class,'xsensenumber')]");
            bool expected = true;
            bool actual;
            actual = Program_Accessor.ApplyAfterRules(afterRules, ClassName, node);
            Assert.AreEqual(expected, actual);
            xDoc.Save(_tf.Output(TestDataName));
            TextFileAssert.AreEqual(_tf.Expected(TestDataName), _tf.Output(TestDataName));
        }

        /// <summary>
        ///A test for ApplyBeforeRule
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void ApplyBeforeRuleTest()
        {
            const string TestDataName = "before.xhtml";
            IDictionary<string, string> beforeRules = new Dictionary<string, string>() { };
            string className = "xsensenumber";
            var xDoc = new XmlDocument();
            xDoc.Load(_tf.Input(TestDataName));
            XmlNode node = xDoc.DocumentElement.SelectSingleNode("//*[contains(@class,'xsensenumber')]");
            bool expected = false;
            bool actual;
            actual = Program_Accessor.ApplyBeforeRule(beforeRules, className, node);
            Assert.AreEqual(expected, actual);
            xDoc.Save(_tf.Output(TestDataName));
            TextFileAssert.AreEqual(_tf.Expected(TestDataName), _tf.Output(TestDataName));
        }

        /// <summary>
        ///A test for ApplyBetweenRules
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void ApplyBetweenRulesTest()
        {
            const string TestDataName = "between.xhtml";
            const string ClassName = "lexref-targets";
            var afterRules = new Dictionary<string, string>() { { ClassName, ", " } };
            var xDoc = new XmlDocument() {XmlResolver = null};
            xDoc.Load(_tf.Input(TestDataName));
            bool expected = true;
            bool actual = false;
            var xpath = string.Format("//*[contains(@class,'{0}')]", ClassName);
            foreach (XmlNode node in xDoc.DocumentElement.SelectNodes(xpath))
            {
                actual |= Program_Accessor.ApplyBetweenRules(afterRules, ClassName, node);
            }
            Assert.AreEqual(expected, actual);
            xDoc.Save(_tf.Output(TestDataName));
            TextFileAssert.AreEqual(_tf.Expected(TestDataName), _tf.Output(TestDataName));
        }

        /// <summary>
        ///A test for ApplyRules
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void ApplyRulesTest()
        {
            const string TestDataName = "applyRules.xhtml";
            var xDoc = new XmlDocument() { XmlResolver = null };
            xDoc.Load(_tf.Input(TestDataName));
            var sr = new StreamReader(_tf.Input("b.css"));
            string cssData = sr.ReadToEnd();
            sr.Close();
            Program_Accessor.ApplyRules(xDoc, cssData);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName), _tf.Output(TestDataName));
        }

        /// <summary>
        ///A test for ChangeDiv2P
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void ChangeDiv2PTest()
        {
            const string TestDataName = "changeDiv2P.xhtml";
            string xpath = "//xhtml:div[xhtml:span]";
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName));
            xDoc.Load(sr);
            sr.Close();
            var nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            Program_Accessor.ChangeDiv2P(xpath, xDoc, nsmgr);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName), _tf.Output(TestDataName));
        }

        /// <summary>
        ///A test for CollectRules
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void CollectRulesTest()
        {
            var sr = new StreamReader(_tf.Input("b.css"));
            string cssData = sr.ReadToEnd();
            sr.Close();
            string patternDesc = "\n\\.([-a-zA-Z]+):after\\s{\\scontent:\\s\"([^\"]*)";
            var expected = new Dictionary<string, string>() { { "xsensenumber", ") " } };
            Dictionary<string, string> actual;
            actual = Program_Accessor.CollectRules(cssData, patternDesc);
            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(actual.ContainsKey("xsensenumber"));
            Assert.AreEqual(expected["xsensenumber"], actual["xsensenumber"]);
        }

        /// <summary>
        ///A test for GetData
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void GetDataTest()
        {
            string fullName = _tf.Input("b.css");
            var sr = new StreamReader(fullName);
            string expected = sr.ReadToEnd();
            sr.Close();
            string actual;
            actual = Program_Accessor.GetData(fullName);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for InsertStyles
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void InsertStylesTest()
        {
            const string TestDataName = "InsertStyles";
            var cssSr = new StreamReader(_tf.Input(TestDataName + ".css"));
            string cssData = cssSr.ReadToEnd();
            cssSr.Close();
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName + ".xhtml"));
            xDoc.Load(sr);
            sr.Close();
            var nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            Program_Accessor.InsertStyles(cssData, xDoc, nsmgr);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName + ".xhtml"));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName + ".xhtml"), _tf.Output(TestDataName + ".xhtml"));
        }

        /// <summary>
        ///A test for InsertStyles border property
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void InsertStylesBorderTest()
        {
            const string TestDataName = "InsertStylesBorder";
            var cssSr = new StreamReader(_tf.Input(TestDataName + ".css"));
            string cssData = cssSr.ReadToEnd();
            cssSr.Close();
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName + ".xhtml"));
            xDoc.Load(sr);
            sr.Close();
            var nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            Program_Accessor.InsertStyles(cssData, xDoc, nsmgr);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName + ".xhtml"));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName + ".xhtml"), _tf.Output(TestDataName + ".xhtml"));
        }

        /// <summary>
        ///A test for InsertTitle
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void InsertTitleTest()
        {
            const string TestDataName = "InsertTitle";
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName + ".xhtml"));
            xDoc.Load(sr);
            sr.Close();
            var nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            string title = "My Title";
            Program_Accessor.InsertTitle(xDoc, nsmgr, title);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName + ".xhtml"));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName + ".xhtml"), _tf.Output(TestDataName + ".xhtml"));
        }

        /// <summary>
        ///A test for UpdateImgSrc
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void UpdateImgSrcTest()
        {
            const string TestDataName = "UpdateImgSrc";
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName + ".xhtml"));
            xDoc.Load(sr);
            sr.Close();
            var nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            Program_Accessor.UpdateImgSrc(xDoc, nsmgr);
            var sw = XmlTextWriter.Create(_tf.Output(TestDataName + ".xhtml"));
            xDoc.Save(sw);
            sw.Close();
            TextFileAssert.AreEqual(_tf.Expected(TestDataName + ".xhtml"), _tf.Output(TestDataName + ".xhtml"));
        }

        /// <summary>
        ///A test for WriteXhtml
        ///</summary>
        [TestMethod()]
        [DeploymentItem("msWordExport.exe")]
        public void WriteXhtmlTest()
        {
            const string TestDataName = "WriteXhtml";
            string output = _tf.Output(TestDataName);
            var xDoc = new XmlDocument() { XmlResolver = null };
            var sr = new StreamReader(_tf.Input(TestDataName + ".xhtml"));
            xDoc.Load(sr);
            sr.Close();
            Program_Accessor.WriteXhtml(output, xDoc);
            var actualSr = new StreamReader(_tf.Output(TestDataName));
            var actual = actualSr.ReadLine();
            actualSr.Close();
            Assert.AreEqual("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.1//EN\" \"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd\"[]><html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"utf-8\" lang=\"utf-8\"><!--", actual);
        }
    }
}
