﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Moq;

namespace RubberduckTests.Mocks
{
    static class MockFactory
    {
        internal static Mock<Window> CreateWindowMock()
        {
            var window = new Mock<Window>();
            window.SetupProperty(w => w.Visible, false);
            window.SetupGet(w => w.LinkedWindows).Returns((LinkedWindows) null);
            window.SetupProperty(w => w.Height);
            window.SetupProperty(w => w.Width);

            return window;
        }

        internal static Mock<VBE> CreateVbeMock(Windows windows)
        {
            var vbe = new Mock<VBE>();
            vbe.Setup(v => v.Windows).Returns(windows);

            return vbe;
        }

        internal static Mock<VBE> CreateVbeMock(Windows windows, VBProjects projects)
        {
            var vbe = CreateVbeMock(windows);
            vbe.SetupGet(v => v.VBProjects).Returns(projects);

            return vbe;
        }

        internal static Mock<CodeModule> CreateCodeModuleMock(string code)
        {
            var lineCount = code.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Length;

            var codeModule = new Mock<CodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(lineCount);
            codeModule.SetupGet(c => c.get_Lines(1, lineCount)).Returns(code);
            return codeModule;
        }

        internal static Mock<VBComponent> CreateComponentMock(string name, CodeModule codeModule, vbext_ComponentType componentType)
        {
            var component = new Mock<VBComponent>();
            component.SetupProperty(c => c.Name, name);
            component.SetupGet(c => c.CodeModule).Returns(codeModule);
            component.SetupGet(c => c.Type).Returns(componentType);
            return component;
        }

        internal static Mock<VBComponents> CreateComponentsMock(List<VBComponent> componentList, VBProject project)
        {
            var components = new Mock<VBComponents>();
            components.Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());
            components.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());
            components.SetupGet(c => c.Parent).Returns(project);

            return components;
        }

        internal static Mock<VBProject> CreateProjectMock(string name, vbext_ProjectProtection protectionLevel)
        {
            var project = new Mock<VBProject>();
            project.SetupProperty(p => p.Name, name);
            project.SetupGet(p => p.Protection).Returns(protectionLevel);
            return project;
        }

        internal static Mock<VBProjects> CreateProjectsMock(List<VBProject> projectList, VBProject project)
        {
            var projects = new Mock<VBProjects>();
            projects.Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());
            projects.As<IEnumerable>().Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());

            return projects;
        }

        //internal static Mock<VBProjects> CreateProjectsMock(List<VBProject> projectList, VBProject project, VBComponents components)
        //{
        //    CreateProjectsMock(projectList, project);
        //    project.SetupGet(p => p.VBComponents).Returns(components.Object);
        //    return projects;
        //}
    }
}