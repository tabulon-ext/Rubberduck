﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class ReferenceResolveRunner : ReferenceResolveRunnerBase
    {
        private const int _maxDegreeOfReferenceResolverParallelism = -1;

        public ReferenceResolveRunner(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IModuleToModuleReferenceManager moduleToModuleReferenceManager,
            IReferenceRemover referenceRemover,
            IDocumentModuleSuperTypeNamesProvider documentModuleSuperTypeNamesProvider) 
        :base(state, 
            parserStateManager, 
            moduleToModuleReferenceManager,
            referenceRemover,
            documentModuleSuperTypeNamesProvider)
        {}


        protected override void ResolveReferences(ICollection<KeyValuePair<QualifiedModuleName, IParseTree>> toResolve, CancellationToken token)
        {
            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceResolverParallelism;

            try
            {
                Parallel.ForEach(toResolve, options,
                    kvp => ResolveReferences(_state.DeclarationFinder, kvp.Key, kvp.Value, token)
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    //This rethrows the exception with the original stack trace.
                    ExceptionDispatchInfo.Capture(exception.InnerException ?? exception).Throw();
                }

                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }

        protected override void AddModuleToModuleReferences(DeclarationFinder finder, CancellationToken token)
        {
            var options = new ParallelOptions
            {
                CancellationToken = token,
                MaxDegreeOfParallelism = _maxDegreeOfReferenceResolverParallelism
            };

            var allModules = finder.AllModules;

            try
            {
                Parallel.ForEach(allModules, options,
                    referencedModule => AddModuleToModuleReferences(finder, referencedModule)
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    //This rethrows the exception with the original stack trace.
                    ExceptionDispatchInfo.Capture(exception.InnerException ?? exception).Throw();
                }

                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }
    }
}
