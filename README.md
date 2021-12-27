# Alternative Installer for X-OPC

Upon request, I share these PowerShell scripts that I use as an alternative way to install and uninstall X-OPC services quickly, as the official installer does a lot of unnecessary steps and offers no GUIless interface. I have no affiliation with HIMA, nor do I take responsibility for compatibility issues or issues, in general, that might arise by using these tools.

![Install X-OPC services](/screenshots/gui.png)

I have found this solution to work in a production-ready environment for X-OPC version `X_OPC V5.2.1204` for OPC DA servers deployed with `SILWorX version 11`.
You would need to download the official binary from HIMA directly as I don't want to share its contents within this repository. This step is required to make this tool work.

#### Necessary steps:

1. Run the binary for X-OPC and install all the required dependencies.
2. Create a zip archive containing EvtLogMsg.dll, Srv_Name.dll, and X-OPC.exe. You will find the binary and the DLL's in 'C:\Program Files (x86)\HIMA\X-OPC. These files are generic
3. Copy the archive to `C:\Temp\hima_service.zip`
4. Uninstall the X-OPC application like you would any other with the Windows OS
5. Run `xopc.ps1` for a GUI where you can specify parameters in YAML. Note that it will try to install `powershell-yaml` the first time around. You would need to have internet access during that time. There are plenty of guides on installing PowerShell dependencies without internet
6. `install` with `uninstall` are interchangeable. As the name suggests, it will either install or uninstall your configuration. I don't recommend editing the file's content within that window, as the data will be lost if an error arises.
7. Run `service.ps1` for the equivalent without GUI. The parameters are pretty self-explanatory, and I suggest looking at the file
