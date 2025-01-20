using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    /// <summary>
    /// An inspection that flags invalid annotation comments.
    /// </summary>
    internal abstract class InvalidAnnotationInspectionBase : InspectionBase
    {
        protected InvalidAnnotationInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected QualifiedContext Context(IParseTreeAnnotation pta) =>
            new QualifiedContext(pta.QualifiedSelection.QualifiedName, pta.Context);

        protected sealed override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
.Where(module => module != null)
.SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder));
        }

        protected IInspectionResult InspectionResult(IParseTreeAnnotation pta) =>
            new QualifiedContextInspectionResult(this, ResultDescription(pta), Context(pta));

        /// <summary>
        /// Gets all invalid annotations covered by this inspection.
        /// </summary>
        /// <param name="annotations">All user code annotations.</param>
        /// <param name="userDeclarations">All user declarations.</param>
        /// <param name="identifierReferences">All identifier references in user code.</param>
        /// <returns></returns>
        protected abstract IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences);

        /// <summary>
        /// Gets an annotation-specific description for an inspection result.
        /// </summary>
        /// <param name="pta">The invalid annotation.</param>
        /// <returns></returns>
        protected abstract string ResultDescription(IParseTreeAnnotation pta);

        protected sealed override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var annotations = finder.FindAnnotations(module);
            var userDeclarations = finder.Members(module).ToList();
            var identifierReferences = finder.IdentifierReferences(module).ToList();

            var invalidAnnotations = GetInvalidAnnotations(annotations, userDeclarations, identifierReferences);
            return invalidAnnotations.Select(InspectionResult).ToList();
        }
    }
}