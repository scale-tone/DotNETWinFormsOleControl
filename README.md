# DotNETWinFormsOleControl

A minimalistic example of a .NET8-based WinForms custom control, that can also be used as an OLE Control in Visual Basic for Applications.

Demonstrates how to scaffold everything, so that it works.

# Prerequisites

Visual Studio 2022 **with C++ Build tools (MSVC v143)** installed.

## How to use

1. Open the solution in Visual Studio.
2. Build the [TypeLib](TypeLib) project. In the `./tlb` folder it will produce a .TLB-file, which is needed by the [Control](Control) project.
3. Build the [Control](Control) project. In its `/bin/Debug/net8.0-windows` folder it will produce the custom control's DLL (`Control.dll`) and also a custom COM activation host library, aka "shim" (`Control.comhost.dll`). At this point you should be able to open that `Control.comhost.dll` file with OleView32 and see the typelib injected into it.
4. Register the *shim* file: `regsvr32 Control.comhost.dll`. You will need an elevated command prompt for that.
5. Open VBA editor in Word or Excel and create an empty form. You should now see the `DotNETWinFormsCustomControlLib.DotNETWinFormsCustomControl` control in the list of "additional controls":

    <img width="500px" src="https://github.com/scale-tone/DotNETWinFormsOleControl/assets/5447190/7f17ef2b-771e-4169-ba6b-dd0977dab6e4"/>

6. Add the onto the form and observe the control working:

    <img width="500px" src="https://github.com/scale-tone/DotNETWinFormsOleControl/assets/5447190/4ee1b8aa-3344-45dd-9d58-8f1b170d40ec"/>


