using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies parameterless 'Range.Cells' member calls.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// Range.Cells is a parameterized Property Get procedure that accepts RowIndex and ColumnIndex parameters, both optional
    /// to avoid requiring either when only one needs to be supplied. If no parameters are provided, 
    /// Cells simply returns a reference to the parent Range object, making a parameterless call entirely redundant.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print Sheet1.Range("A1").Cells.Address
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print Sheet1.Range("A1").Address
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal sealed class ParameterlessCellsInspection : IdentifierReferenceInspectionBase
    {
        public ParameterlessCellsInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var excel = finder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel is null)
            {
                yield break;
            }

            var range = finder.Classes.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Range" && item.ProjectId == excel.ProjectId);
            if (range is null)
            {
                yield break;
            }

            var cells = finder.Members(range).SingleOrDefault(item => item.IdentifierName == "Cells" && item.DeclarationType == DeclarationType.PropertyGet);
            if (cells is null)
            {
                yield break;
            }

            foreach (var reference in cells.References.Where(reference => IsResultReference(reference, finder)))
            {
                yield return InspectionResult(reference, finder);
            }
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var memberAccess = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
            var memberArgs = memberAccess?.GetAncestor<VBAParser.IndexExprContext>()?.argumentList()?.argument();

            return memberAccess is VBAParser.MemberAccessExprContext && (memberArgs?.Length ?? 0) == 0;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return InspectionResults.ResourceManager.GetString(nameof(ParameterlessCellsInspection), CultureInfo.CurrentUICulture);
        }
    }
}
