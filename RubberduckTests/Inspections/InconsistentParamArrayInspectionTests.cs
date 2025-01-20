using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class InconsistentParamArrayBaseInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void InconsistentParamArrayBase_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Public Sub DoSomething(ParamArray Values)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WithoutOptionBase_NoResult()
        {
            const string inputCode =
                @"Public Sub DoSomething(ParamArray Values)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());

        }

        [Test]
        [Category("Inspections")]
        public void NonParamArrayParameter_NoResult()
        {
            const string inputCode =
                @"Option Base 1
Public Sub DoSomething(ByRef Values() As Variant)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new InconsistentParamArrayBaseInspection(state);
        }
    }

    [TestFixture]
    public class InconsistentArrayBaseInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void InconsistentArrayBase_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Public Sub DoSomething()
    Dim Values As Variant
    Values = VBA.Array(42)
End Sub";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WithoutOptionBase_NoResult()
        {
            const string inputCode =
                @"Public Sub DoSomething()
    Dim Values As Variant
    Values = VBA.Array(42)
End Sub";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());

        }

        [Test]
        [Category("Inspections")]
        public void WithoutQualifier_NoResult()
        {
            const string inputCode =
                @"Public Sub DoSomething()
    Dim Values As Variant
    Values = Array(42)
End Sub";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());

        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode)
        {
            return InspectionResultsForModules(("TestModule1", inputCode, ComponentType.StandardModule), ReferenceLibrary.VBA);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new InconsistentArrayBaseInspection(state);
        }
    }
}
