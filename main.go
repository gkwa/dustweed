package main

import (
	"fmt"
	"os"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	// Initialize COM library
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	if err != nil {
		fmt.Println("Failed to initialize COM:", err)
		os.Exit(1)
	}
	defer ole.CoUninitialize()

	// Create a new Shell object
	unknown, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		fmt.Println("Failed to create WScript.Shell object:", err)
		os.Exit(1)
	}
	shell, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Println("Failed to get IDispatch:", err)
		os.Exit(1)
	}
	defer shell.Release()

	// Get the Desktop folder
	csidlDesktop := 0x0000
	shellDesktop, err := oleutil.CallMethod(shell, "Namespace", csidlDesktop)
	if err != nil {
		fmt.Println("Failed to get Desktop folder:", err)
		os.Exit(1)
	}
	folder := shellDesktop.ToIDispatch()
	defer folder.Release()

	// Create the shortcut
	csidlWindows := 0x0024
	shortcut, err := oleutil.CallMethod(folder, "ParseName", csidlWindows)
	if err != nil {
		fmt.Println("Failed to create shortcut:", err)
		os.Exit(1)
	}

	_, err = oleutil.CallMethod(shortcut.ToIDispatch(), "InvokeVerb", "PowerShell")
	if err != nil {
		fmt.Println("Failed to create PowerShell shortcut:", err)
		os.Exit(1)
	}

	fmt.Println("PowerShell shortcut created successfully.")
}
