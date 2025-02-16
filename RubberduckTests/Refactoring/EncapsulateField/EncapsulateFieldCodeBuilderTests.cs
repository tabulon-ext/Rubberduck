using NUnit.Framework;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using System;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldCodeBuilderTests
    {
        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldCodeBuilder))]
        public void BuildPropertyBlock_VariantGet(string accessibility)
        {
            var attrSet = new PropertyAttributeSet()
            {
                PropertyName = "NewProperty",
                BackingField = "backingField",
                AsTypeName = "Variant",
            };

            var results = GeneratePropertyBlocks($"{accessibility} xxx As Variant", "xxx", attrSet);
            var actualLines = results.Get.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            Assert.AreEqual(7, actualLines.Count);
            Assert.AreEqual(actualLines[0].Trim(), $"{accessibility} Property Get {attrSet.PropertyName}() As {attrSet.AsTypeName}");
            Assert.AreEqual(actualLines[1].Trim(), $"If IsObject({attrSet.BackingField}) Then");
            Assert.AreEqual(actualLines[2].Trim(), $"Set {attrSet.PropertyName} = {attrSet.BackingField}");
            Assert.AreEqual(actualLines[3].Trim(), "Else");
            Assert.AreEqual(actualLines[4].Trim(), $"{attrSet.PropertyName} = {attrSet.BackingField}");
            Assert.AreEqual(actualLines[5].Trim(), "End If");
            Assert.AreEqual(actualLines[6].Trim(), "End Property");
        }

        [TestCase("Variant", "Public")]
        [TestCase("Long", "Public")]
        [TestCase("String", "Public")]
        [TestCase("Variant", "Private")]
        [TestCase("Long", "Private")]
        [TestCase("String", "Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldCodeBuilder))]
        public void BuildPropertyBlock_Let(string asTypeName, string accessibility)
        {
            var attrSet = new PropertyAttributeSet()
            {
                PropertyName = "NewProperty",
                BackingField = "backingField",
                RHSParameterIdentifier = "RHS",
                GeneratePropertyLet = true,
                AsTypeName = asTypeName,
            };

            var results = GeneratePropertyBlocks($"{accessibility} xxx As {asTypeName}", "xxx", attrSet);
            var actualLines = results.Let.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            Assert.AreEqual(3, actualLines.Count);
            Assert.AreEqual(actualLines[0].Trim(), $"{accessibility} Property Let {attrSet.PropertyName}(ByVal {attrSet.RHSParameterIdentifier} As {asTypeName})");
            Assert.AreEqual(actualLines[1].Trim(), $"{attrSet.BackingField} = {attrSet.RHSParameterIdentifier}");
            Assert.AreEqual(actualLines[2].Trim(), "End Property");
        }

        [TestCase("Variant", "Public")]
        [TestCase("Long", "Public")]
        [TestCase("String", "Public")]
        [TestCase("Variant", "Private")]
        [TestCase("Long", "Private")]
        [TestCase("String", "Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldCodeBuilder))]
        public void BuildPropertyBlock_Set(string asTypeName, string accessibility)
        {
            var attrSet = new PropertyAttributeSet()
            {
                PropertyName = "NewProperty",
                BackingField = "backingField",
                RHSParameterIdentifier = "RHS",
                GeneratePropertySet = true,
                AsTypeName = asTypeName,
            };

            var results = GeneratePropertyBlocks($"{accessibility} xxx As {asTypeName}", "xxx", attrSet);
            var actualLines = results.Set.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            Assert.AreEqual(3, actualLines.Count);
            Assert.AreEqual(actualLines[0].Trim(), $"{accessibility} Property Set {attrSet.PropertyName}(ByVal {attrSet.RHSParameterIdentifier} As {asTypeName})");
            Assert.AreEqual(actualLines[1].Trim(), $"Set {attrSet.BackingField} = {attrSet.RHSParameterIdentifier}");
            Assert.AreEqual(actualLines[2].Trim(), "End Property");
        }

        private (string Get, string Let, string Set) GeneratePropertyBlocks(string code, string prototypeIdentifier, PropertyAttributeSet attrSet)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var encapsulateTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals(prototypeIdentifier));

                attrSet.Declaration = encapsulateTarget;
                var resolver = EncapsulateFieldTestSupport.GetResolver(state);

                return resolver.Resolve<IEncapsulateFieldCodeBuilder>()
                    .BuildPropertyBlocks(attrSet);
            }
        }
    }
}
