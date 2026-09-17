//go:build windows

package main

import (
	"testing"

	"golang.org/x/sys/windows"
)

func TestConfigureBackgroundCommandHidesConsoleWindow(t *testing.T) {
	command := newBackgroundCommand("powershell.exe", "-NoProfile", "-Command", "exit 0")

	if command.SysProcAttr == nil {
		t.Fatal("SysProcAttr is nil; helper process can create a visible console window")
	}
	if !command.SysProcAttr.HideWindow {
		t.Fatal("HideWindow is false; helper process can create a visible console window")
	}
	if command.SysProcAttr.CreationFlags&windows.CREATE_NO_WINDOW == 0 {
		t.Fatal("CREATE_NO_WINDOW is missing; helper process can create a visible console window")
	}
}
