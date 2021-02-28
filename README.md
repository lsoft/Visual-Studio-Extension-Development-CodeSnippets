# Visual-Studio-Extension-Development-CodeSnippets
Some useful links and pieces of code in the scope of Visual Studio Extension (vsix) development



# Unsorted

## Visual Studio's KnownMonikers

The images contained in Visual Studio's IVsImageService: http://glyphlist.azurewebsites.net/knownmonikers/

## Wait for solution is fully loaded

```csharp
                var solution = AsyncPackage.GetGlobalService(typeof(SVsSolution)) as IVsSolution;

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                object asm;
                solution.GetProperty((int)__VSPROPID4.VSPROPID_IsSolutionFullyLoaded, out asm);

                if (asm is bool)
                {
                    if ((bool)asm)
                    {
                        _isSolutionFullyLoaded = true;
                        return;
                    }
                }

```