using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags Rubberduck annotations used in a component type that is incompatible with that annotation.
    /// </summary>
    /// <why>
    /// Some annotations can only be used in a specific type of module; others cannot be used in certain types of modules.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@PredeclaredId  'this annotation is illegal in a standard module
    /// Option Explicit
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@PredeclaredId  'this annotation works fine in a class module
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class AnnotationInIncompatibleComponentTypeInspection : InvalidAnnotationInspectionBase
    {
        public AnnotationInIncompatibleComponentTypeInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            foreach (var pta in annotations)
            {
                var annotation = pta.Annotation;
                var componentType = pta.QualifiedSelection.QualifiedName.ComponentType;
                if (annotation.RequiredComponentType.HasValue && annotation.RequiredComponentType != componentType
                       || annotation.IncompatibleComponentTypes.Contains(componentType))
                {
                    yield return pta;
                }
            }

            yield break;
        }

        protected override string ResultDescription(IParseTreeAnnotation pta)
        {
            if (pta.Annotation.RequiredComponentType.HasValue)
            {
                return string.Format(InspectionResults.InvalidAnnotationInspection_NotInRequiredComponentType,
                    pta.Annotation.Name, // annotation...
                    pta.QualifiedSelection.QualifiedName.ComponentType,  // is used in a...
                    pta.Annotation.RequiredComponentType); // but is only valid in a...
            }
            else
            {
                return string.Format(InspectionResults.InvalidAnnotationInspection_IncompatibleComponentType,
                    pta.Annotation.Name, // annotation...
                    pta.QualifiedSelection.QualifiedName.ComponentType); // cannot be used in a...
            }
        }
    }
}