using Rubberduck.CodeAnalysis.Inspections.Abstract;
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
    /// Flags comments that parsed like Rubberduck annotations, but were not recognized as such.
    /// </summary>
    /// <why>
    /// Other add-ins may support similar-looking annotations that Rubberduck does not recognize; this inspection can be used to spot a typo in Rubberduck annotations.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Param "Value", "The value to print."  : Rubberduck does not define a @Param annotation
    /// Public Sub Test(ByVal Value As Long)
    ///     Debug.Print Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@Description "Prints the specified value." : Rubberduck defines a @Description annotation
    /// Public Sub Test(ByVal Value As Long)
    ///     Debug.Print Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UnrecognizedAnnotationInspection : InvalidAnnotationInspectionBase
    {
        public UnrecognizedAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            return annotations.Where(pta => pta.Annotation is NotRecognizedAnnotation).ToList();
        }

        protected override string ResultDescription(IParseTreeAnnotation pta) =>
            string.Format(InspectionResults.ResourceManager.GetString(nameof(UnrecognizedAnnotationInspection), CultureInfo.CurrentUICulture), pta.Context.GetText());
    }
}