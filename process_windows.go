//go:build windows

package main

import (
	"os/exec"
	"syscall"

	"golang.org/x/sys/windows"
)

// configureBackgroundCommand keeps helper processes invisible when the app is
// built as a Windows GUI executable. Without these flags, PowerShell briefly
// creates a console window whenever Outlook status is checked or an import starts.
func configureBackgroundCommand(command *exec.Cmd) {
	command.SysProcAttr = &syscall.SysProcAttr{
		CreationFlags: windows.CREATE_NO_WINDOW,
		HideWindow:    true,
	}
}
