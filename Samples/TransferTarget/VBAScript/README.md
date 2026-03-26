# VBAScript — Send Excel Sheet to an app via TransferTarget API

This folder contains a VBA macro that adds a **"Send to Email"** button to any Excel worksheet. Clicking the button reads the active sheet's contents and sends them to the configured app (such as **Outlook (New)**) using the Windows [TransferTarget API](https://learn.microsoft.com/en-us/windows/apps/develop/share-and-transfer/transfer-target).

Because VBA cannot call WinRT APIs directly, a small C# COM bridge (`TransferTargetBridge`) acts as the intermediary.

## Folder structure

```
VBAScript/
├── README.md                       ← This file
├── SendToEmail.ba  s               ← VBA module source (the macro code)
├── BuildAddin.vbs                  ← Script that packages the .bas into an Excel Add-in (.xlam)
└── TransferTargetBridge/           ← C# COM server that exposes the TransferTarget API to VBA
    ├── TransferTargetBridge.sln
    ├── TransferTargetBridge.csproj
    └── TransferTargetHelper.cs     ← COM-visible class with ShareEmailToApp()
```

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 11 Build **26100.7015** or later (TransferTarget API support) |
| **SDK** | Windows SDK **10.0.26100.0** or later |
| **.NET** | .NET 10 SDK (the bridge targets `net10.0-windows10.0.26100.0`) |
| **Visual Studio** | 2026 or later with the **.NET desktop development** workload |
| **Excel** | Microsoft Excel (desktop) with VBA macro support enabled |
| An app that supports Share Target | Such as **Outlook (New)** |

## Setup (step by step)

### Step 1 — Build and register the COM bridge

The COM bridge is a C# class library that exposes the TransferTarget API as a COM object so VBA can call it.

1. Open a **Developer Command Prompt** (or any terminal with `dotnet` and `regsvr32` available).

2. Navigate to the bridge project folder:

   ```
   cd VBAScript\TransferTargetBridge
   ```

3. Build the project:

   ```
   dotnet build -c Debug -r win-x64
   ```

4. After a successful build the output folder will contain `TransferTargetBridge.comhost.dll`. Register it with COM by running an **elevated (Administrator)** command prompt:

   ```
   regsvr32 "bin\Debug\net10.0-windows10.0.26100.0\win-x64\TransferTargetBridge.comhost.dll"
   ```

   You should see a message saying registration succeeded.

> **Tip:** To unregister the COM server, run `regsvr32 /u` with the same path.

### Step 2 - Customize the VBA macro

In the file `SendToEmail.bas`, make the following changes:

* Specify the program that will be the target of the "Send to Email" button
  by setting the `APPID` variable to the AppuserModelId of the app you want to use.
  The sample comes preconfigured to use **Outlook (New)**, but you can substitute
  any other program.

> **Note:** The target app must support the Share contract and be able to receive `text` or `HTML` data formats.
> You can experiment with the accompanying TransferTarget sample to see which target apps support which data formats.

* Optionally specify a path to a bitmap file (such as a png) that will be used
  as an icon in the custom button. If you do not specify one, then the custom button
  will not have an icon.

### Step 3 — Import the VBA macro into Excel

You have **two options**:

#### Option A — Import the `.bas` file manually

1. Open any Excel workbook (or create a new one).
2. Press **Alt + F11** to open the VBA Editor.
3. Go to **File → Import File...** and select `SendToEmail.bas`.
4. Close the VBA Editor.
5. Save the workbook as **Excel Macro-Enabled Workbook (.xlsm)**.

#### Option B — Build an Excel Add-in automatically

The `BuildAddin.vbs` script creates an `.xlam` add-in file so the button is available in *every* workbook.

1. **Enable programmatic access to VBA** (required for the script to import the `.bas` file):
   - In Excel, go to **File → Options → Trust Center → Trust Center Settings → Macro Settings**.
   - Check **Disable VBA macros with notification**.
   - Click **OK**.

2. Run the script from a command prompt (it automatically finds `SendToEmail.bas` in the same directory):

   ```
   cscript BuildAddin.vbs
   ```

   This creates `SendToEmail.xlam` in `%APPDATA%\Microsoft\AddIns\`.

3. In Excel, go to **File → Options → Add-Ins → Manage: Excel Add-ins → Go...** and check **SendToEmail** to enable it.

## Usage

1. Open an Excel workbook with some data in the active sheet.

2. Run the macro `AddSendToEmailButton`:
   - Press **Alt + F8**, select **AddSendToEmailButton**, and click **Run**.
   - This places a button in the top-left area of the active sheet.

3. Click the **"Send to Outlook"** button.

4. The macro will:
   1. Read the active sheet's contents and convert them to both plain text and HTML (preserving formatting).
   2. Create a COM bridge instance (`TransferTargetBridge.Helper`).
   3. Use a `TransferTargetWatcher` to discover a specific transfer target.
   4. Transfer the data to that target. If the target is an email program, it will open a new email compose window with the sheet contents.

5. A success or error message will appear when the operation completes.

## Troubleshooting

| Problem | Solution |
|---|---|
| **"Could not create TransferTargetBridge.Helper"** | The COM DLL is not registered. Run `regsvr32` as admin (see Step 1). |
| **"TransferTargetWatcher is not available"** | The TransferTarget feature requires Windows 11 Build 26100.7015 or later. |
| **"No matching target app found"** | Make sure the app specified by the APPID is installed, and can receive text and HTML content. You may need to do some setup in the app before it will start accepting data. (For example, you may need to set up an email account.) |
| **Macro security blocks execution** | In Excel: **File → Options → Trust Center → Macro Settings** → select **Disable with notification** and click Enable when prompted for the specific document/add-in. |
| **`BuildAddin.vbs` fails with "Programmatic access" error** | Enable **Trust access to the VBA project object model** in Excel Trust Center settings (see Step 2, Option B) temporarily. |
