using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ParameterlessCellsInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]

        public void ParameterlessCells_ReturnsResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells.Address
End Sub
";
            Assert.AreEqual(1, InspectionResultsFor(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void CellsWithEmptyArgsList_ReturnsResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells().Address
End Sub
";
            Assert.AreEqual(1, InspectionResultsFor(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void CellsWithRowIndexArgument_NoResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells(42).Address
End Sub
";
            Assert.AreEqual(0, InspectionResultsFor(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void CellsWithNamedRowIndexArgument_NoResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells(RowIndex:=42).Address
End Sub
";
            Assert.AreEqual(0, InspectionResultsFor(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void CellsWithNamedColumnIndexArgument_NoResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells(ColumnIndex:=42).Address
End Sub
";
            Assert.AreEqual(0, InspectionResultsFor(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void CellsWithBothArguments_NoResult()
        {
            const string inputCode = @"Option Explicit
Private Sub DoSomething()
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Debug.Print Sheet.Range(""A1"").Cells(42, 1).Address
End Sub
";
            Assert.AreEqual(0, InspectionResultsFor(inputCode).Count());
        }

        private IEnumerable<IInspectionResult> InspectionResultsFor(string inputCode) =>
            InspectionResultsForModules(
                ("TestModule1", inputCode, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule),
                new[] { Mocks.ReferenceLibrary.VBA, Mocks.ReferenceLibrary.Excel });

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ParameterlessCellsInspection(state);
        }
    }
}
