# Visual Studio Extension Development Tips && CodeSnippets
Some useful links and pieces of code in the scope of Visual Studio Extension (vsix) development.


# General

## Extensions for the Visual Studio gallery

https://marketplace.visualstudio.com/

## Get Started

https://docs.microsoft.com/en-us/visualstudio/extensibility/?view=vs-2019&amp%3BWT.mc_id=docs-twitter-jamont

## List of available services for extension author

https://docs.microsoft.com/en-us/visualstudio/extensibility/internals/list-of-available-services?view=vs-2019

## Chat

https://gitter.im/Microsoft/extendvs

## A place were questions lives

https://docs.microsoft.com/en-us/answers/topics/vs-extensions.html

# Youtube channel

https://www.youtube.com/playlist?list=PLReL099Y5nRdG2n1PrY_tbCsUznoYvqkS


# Code snippets

## MessageBox

```csharp
VsShellUtilities.ShowMessageBox(
    this.package,
    "Cannot parse configuration XML file. Please verify it and try again.",
    "Configuration error",
    OLEMSGICON.OLEMSGICON_CRITICAL,
    OLEMSGBUTTON.OLEMSGBUTTON_OK,
    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST
    );
```

Also, [this video](https://www.youtube.com/watch?v=kqcOg0b_XhA&list=PLReL099Y5nRdG2n1PrY_tbCsUznoYvqkS&index=2) is useful.

## Wait for solution is fully loaded

Run periodically:

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

## Visual Studio events

Use 

```csharp
DTE.DTEEvents
```

For example `EnvDTE.DTE.DTEEvents.OnBeginShutdown` should be used to determine the moment of VS is started to close. You need to take into account [this link](https://social.msdn.microsoft.com/Forums/en-US/eb6cc3eb-422a-48b1-86da-7a81d3edbddc/events-not-captured-afte-a-window-is-opened?forum=vsx) - `Your solution events aren't firing because the objects are getting collected`.

## Open file in Visual Studio

```csharp
dte.ItemOperations.OpenFile(documentFullPath, "{" + VSConstants.LOGVIEWID_Code + "}");
```

## Open file in Visual Studio and navigate to the required piece of code

```csharp
public static void OpenAndNavigate(
    string documentFullPath,
    int startLine,
    int startColumn,
    int endLine,
    int endColumn
    )
{
    if (documentFullPath == null)
    {
        throw new ArgumentNullException(nameof(documentFullPath));
    }

    ThreadHelper.ThrowIfNotOnUIThread(nameof(VisualStudioHelper.OpenAndNavigate));

    var openDoc = AsyncPackage.GetGlobalService(typeof(IVsUIShellOpenDocument)) as IVsUIShellOpenDocument;

    IVsWindowFrame frame;
    Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp;
    IVsUIHierarchy hier;
    uint itemid;
    Guid logicalView = VSConstants.LOGVIEWID_Code;

    if (ErrorHandler.Failed(
        openDoc.OpenDocumentViaProject(
            documentFullPath,
            ref logicalView,
            out sp,
            out hier,
            out itemid,
            out frame)
        )
        || frame == null)
    {
        return;
    }

    object docData;
    frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocData, out docData);

    // Get the VsTextBuffer  
    var buffer = docData as VsTextBuffer;
    if (buffer == null)
    {
        IVsTextBufferProvider bufferProvider = docData as IVsTextBufferProvider;
        if (bufferProvider != null)
        {
            IVsTextLines lines;
            ErrorHandler.ThrowOnFailure(
                bufferProvider.GetTextBuffer(out lines)
                );

            buffer = lines as VsTextBuffer;

            Debug.Assert(buffer != null, "IVsTextLines does not implement IVsTextBuffer");

            if (buffer == null)
            {
                return;
            }
        }
    }

    IVsTextManager textManager = Package.GetGlobalService(typeof(VsTextManagerClass)) as IVsTextManager;

    var docViewType = default(Guid);

    textManager.NavigateToLineAndColumn(
        buffer,
        ref docViewType,
        startLine,
        startColumn,
        endLine,
        endColumn
        );
}
```


# Unsorted

## Visual Studio's KnownMonikers

The images contained in Visual Studio's IVsImageService: http://glyphlist.azurewebsites.net/knownmonikers/

Makes it easier to create and maintain .imagemanifest files: https://marketplace.visualstudio.com/items?itemName=MadsKristensen.ImageManifestTools

## Debugging codelens

To debug your implementation of `IAsyncCodeLensDataPointProvider` you need to attach `ServiceHub.HostCLR.x86.exe`, otherwise not breakpoint hit occurs. There are a few these processed, so attach them all!
