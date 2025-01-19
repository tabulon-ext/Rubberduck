using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Globalization;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates non-breaking spaces hidden in identifier names.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// This inspection may accidentally reveal non-breaking spaces in code copied and pasted from a website.
    /// </why>
    /// <remarks>
    /// You may have discovered this inspection by pasting code directly from a web page, which often contains such non-printable characters.
    /// </remarks>
    internal sealed class NonBreakingSpaceIdentifierInspection : DeclarationInspectionBase
    {
        private const string Nbsp = "\u00A0";

        public NonBreakingSpaceIdentifierInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IdentifierName.Contains(Nbsp);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return InspectionResults.ResourceManager.GetString("NonBreakingSpaceIdentifierInspection", CultureInfo.CurrentUICulture).ThunderCodeFormat(declaration.IdentifierName);
        }
    }
}
