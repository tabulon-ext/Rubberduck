using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about inconsistent implicit lower bounds of arrays when 'Option Base 1' is specified.
    /// </summary>
    /// <why>
    /// The base of a ParamArray is always zero; the VBA.Array function, when explicitly qualified with 'VBA.', is also always zero-based.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Base 1 '<~ Implicit array lower bound is 1
    /// Public Sub DoSomething()
    ///     Dim Values As Variant
    ///     
    ///     Values = Array(42)
    ///     Debug.Print LBound(Values) '<~ 1 as per Option Base
    ///     
    ///     Values = VBA.Array(42) '<<< inspection result here
    ///     Debug.Print LBound(Values) '<~ not 1
    /// End Sub
    /// ]]>
    /// </module>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Base 1 '<~ Implicit array lower bound is 1
    /// Public Sub DoSomething(ParamArray Values) '<<< inspection result here
    ///     Debug.Print LBound(Values) '<~ not 1
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim Values As Variant
    ///     
    ///     Values = Array(42)
    ///     Debug.Print LBound(Values) '<~ 0
    ///     
    ///     Values = VBA.Array(42)
    ///     Debug.Print LBound(Values) '<~ also 0
    /// End Sub
    /// ]]>
    /// </module>
    internal sealed class InconsistentArrayBaseInspection : IdentifierReferenceInspectionBase
    {
        public InconsistentArrayBaseInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var parentModule = finder.ModuleDeclaration(reference.QualifiedModuleName);
            var hasOptionBase1 = parentModule.Context.GetDescendent<VBAParser.OptionBaseStmtContext>()?.numberLiteral()?.GetText() == "1";

            if (reference.Declaration is ParameterDeclaration parameter)
            {
                if (parameter.IsParamArray)
                {
                    return hasOptionBase1;
                }
            }

            if (reference.Declaration.ProjectName == "VBA" && reference.Declaration.IdentifierName == "Array")
            {
                if (reference.QualifyingReference?.Declaration.IdentifierName == "VBA")
                {
                    return true;
                }
            }

            return false;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return reference.Declaration is ParameterDeclaration parameter && parameter.IsParamArray
                ? $"Parameter array '{reference.IdentifierName}' is always zero-based"
                : $"Qualified VBA.Array function always returns a zero-based array";
        }
    }
}
