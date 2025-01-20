using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{

    /// <summary>
    /// Flags invalid or misplaced Rubberduck annotation comments.
    /// </summary>
    /// <why>
    /// Rubberduck is correctly parsing an annotation, but that annotation is illegal in that context and couldn't be bound to a code element.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     '@Folder("Module1.DoSomething")
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Module1.DoSomething")
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class InvalidAnnotationInspection : InvalidAnnotationInspectionBase
    {
        public InvalidAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        protected override string ResultDescription(IParseTreeAnnotation pta) =>
            string.Format(InspectionResults.ResourceManager.GetString(nameof(InvalidAnnotationInspection), CultureInfo.CurrentUICulture), pta.Annotation.Name);

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            return GetUnboundAnnotations(annotations, userDeclarations, identifierReferences)
                .Where(pta => !pta.Annotation.Target.HasFlag(AnnotationTarget.General) || pta.AnnotatedLine == null)
                .Concat(AttributeAnnotationsOnDeclarationsNotAllowingAttributes(userDeclarations))
                .ToList();
        }

        private IEnumerable<IParseTreeAnnotation> GetUnboundAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            var boundAnnotationsSelections = userDeclarations
                .SelectMany(declaration => declaration.Annotations)
                .Concat(identifierReferences.SelectMany(reference => reference.Annotations))
                .Select(annotation => annotation.QualifiedSelection)
                .ToHashSet();

            return annotations
                .Where(pta => pta.Annotation.GetType() != typeof(NotRecognizedAnnotation) && !boundAnnotationsSelections.Contains(pta.QualifiedSelection));
        }

        private IEnumerable<IParseTreeAnnotation> AttributeAnnotationsOnDeclarationsNotAllowingAttributes(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.AttributesPassContext == null
                                      && !declaration.DeclarationType.HasFlag(DeclarationType.Module))
                .SelectMany(declaration => declaration.Annotations)
                .Where(pta => pta.Annotation is IAttributeAnnotation);
        }
    }
}